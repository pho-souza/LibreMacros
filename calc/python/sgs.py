# -*- coding: utf-8 -*-
"""Atualiza séries do SGS (Banco Central) no LibreOffice Calc.

Lê os IDs em 'Parametros_SGS' (A2:D...) e grava a tabela em 'Tabela_SGS'.
Colunas de 'Parametros_SGS':
  A = ID da série | B = descrição (opcional) | C = data inicial | D = data final
  E = status por linha (preenchido pela macro: verde = ok, amarelo = aviso, vermelho = erro)
  F = log geral (preenchido pela macro)

Comportamento:
  - sem datas, baixa a série inteira (consulta sem filtro); se a API recusar, refaz em
    janelas de até 10 anos indo do mais recente para o mais antigo e PARA ao passar do início
    da série (o BCB limita consultas por período a 10 anos desde 26/03/2025); intervalos
    informados maiores que 10 anos são divididos em janelas;
  - antes de baixar, uma checagem rápida (último valor) detecta série inexistente em até TIMEOUT_EXISTENCIA_SEGUNDOS;
    a espera por dados é maior (TIMEOUT_SEGUNDOS), porque janelas de série diária podem demorar;
  - respostas pedidas com gzip; na coluna E, cada série mostra nº de requisições, tempo de
    rede e tamanho, para ver onde o tempo foi; 429 (limite da API) respeita o Retry-After;
  - série com erro (inexistente, fora do ar, etc.) é apontada na coluna E; a coluna
    que ela já tinha em 'Tabela_SGS' é mantida (o vínculo é pelo título da coluna,
    ou seja, a descrição da coluna B ou o ID);
  - se nenhuma série for baixada, a tabela não é alterada; se a gravação falhar,
    o conteúdo anterior é restaurado;
  - valor ausente fica em branco; downloads em paralelo, com timeout e nova tentativa;
  - EXECUCAO_ASSINCRONA = True (padrão): baixa em segundo plano, a planilha fica livre e a gravação
    acontece sozinha no fim, com aviso na barra de status (AVISO_FINAL); False espera terminar.

Instalação: copie para a pasta de scripts Python do usuário
(Linux: ~/.config/libreoffice/4/user/Scripts/python/ ; no Flatpak:
~/.var/app/org.libreoffice.LibreOffice/config/libreoffice/4/user/Scripts/python/)
e execute AtualizarSGS por Ferramentas > Macros > Executar macro.
NÃO foi testado dentro do LibreOffice.
"""
import gzip
import json
import socket
import threading
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta

import uno

HOSTS_API = ("api.bcb.gov.br",)
CONEXAO_RAPIDA_SEGUNDOS = 2.0  # espera máxima por CADA endereço do servidor ao conectar (evita travar num endereço fora do ar)
PREFERIR_IPV4 = True          # tenta o IPv4 antes do IPv6 nestes hosts (evita espera quando o IPv6 está bloqueado)
NOME_FONTE = "SGS"
ROTULO_ID = "ID da Série"
ABA_PARAMETROS = "Parametros_SGS"
ABA_DESTINO = "Tabela_SGS"
MAX_LINHAS_PARAMETROS = 200   # linhas lidas em 'Parametros_SGS' (a partir da 2)
MAX_THREADS = 4               # poucos downloads simultâneos, para não ser barrado pela API
EXECUCAO_ASSINCRONA = True    # baixa em segundo plano e deixa a planilha livre; False = espera terminar (planilha travada)
AVISO_FINAL = "BARRA"         # ao terminar em segundo plano: "BARRA" (barra de status), "CAIXA" (janela com OK) ou None (só o log em F2)
AVISO_SEGUNDOS = 8            # quanto tempo a mensagem da barra de status fica visível
LIMITE_EXECUCAO_SEGUNDOS = 900  # uma execução em segundo plano que passe disso é considerada travada e não bloqueia a próxima
TIMEOUT_SEGUNDOS = 30         # espera máxima por resposta DE DADOS (uma janela de série diária pode levar 10-20 s no BCB)
TIMEOUT_EXISTENCIA_SEGUNDOS = 8   # espera máxima pela checagem "a série existe?"; série inexistente não recebe resposta, e esse é o tempo gasto por ela
VERIFICAR_EXISTENCIA = True   # antes de baixar, pede o último valor da série (resposta rápida) para detectar série inexistente
JANELA_ANOS = 3               # tamanho das janelas quando a API recusa o pedido sem filtro (séries diárias); o máximo permitido é 10
JANELA_DIAS = 3650            # < 10 anos, limite da API
DATA_MINIMA = date(1900, 1, 1)  # início do histórico no modo de janelas (só usado no fallback)
EPOCA_CALC = date(1899, 12, 30)
URL_SEM_FILTRO = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.{sid}/dados?formato=json"
URL = URL_SEM_FILTRO + "&dataInicial={ini}&dataFinal={fim}"
URL_ULTIMOS = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.{sid}/dados/ultimos/1?formato=json"

# CellFlags: VALUE(1) + DATETIME(2) + STRING(4) + FORMULA(16) + HARDATTR(32)
LIMPAR = 1 | 2 | 4 | 16 | 32


# ---------------------------------------------------------------- específico do SGS

class ErroHTTP(RuntimeError):
    def __init__(self, codigo, msg):
        RuntimeError.__init__(self, msg)
        self.codigo = codigo


class ErroSemResposta(RuntimeError):
    """O servidor aceitou a conexão mas não respondeu (típico de série inexistente). Não adianta repetir."""


def _normalizar_id(valor):
    sid = str(int(valor)) if isinstance(valor, float) else str(valor).strip()
    if not sid.isdigit():
        raise ValueError("ID inválido '%s' (use só números)" % sid)
    return sid


def _serial(d):
    return float((d - EPOCA_CALC).days)


def _janelas(ini, fim):
    """Janelas de até 10 anos, do mais antigo para o mais recente."""
    atual = ini
    while atual <= fim:
        fim_janela = min(atual + timedelta(days=JANELA_DIAS - 1), fim)
        yield atual, fim_janela
        atual = fim_janela + timedelta(days=1)


def _janelas_retroativas(fim, ini_min, dias=None):
    """Janelas do mais recente para o mais antigo (padrão: JANELA_ANOS anos cada)."""
    dias = dias or min(JANELA_ANOS, 10) * 365
    atual_fim = fim
    while atual_fim >= ini_min:
        atual_ini = max(ini_min, atual_fim - timedelta(days=dias - 1))
        yield atual_ini, atual_fim
        atual_fim = atual_ini - timedelta(days=1)


def _espera_429(erro, tentativa):
    try:
        return min(30.0, float(erro.headers.get("Retry-After")))
    except (TypeError, ValueError, AttributeError):
        return 3.0 * (tentativa + 1)


def _get_json(url, tentativas=3, stats=None, timeout=None, dica=""):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (macro SGS LibreOffice)",
                                               "Accept": "application/json", "Accept-Encoding": "gzip"})
    st = stats if stats is not None else {"req": 0, "rede": 0.0, "leitura": 0.0, "bytes": 0, "extras": 0}
    limite = timeout or TIMEOUT_SEGUNDOS
    ultimo = None
    for t in range(tentativas):
        t0 = time.time()
        espera = 0.5 * (2 ** t)
        try:
            with urllib.request.urlopen(req, timeout=limite) as resp:
                bruto = resp.read()
                comprimido = (resp.headers.get("Content-Encoding") or "").lower() == "gzip"
            st["req"] += 1
            st["rede"] += time.time() - t0
            st["bytes"] += len(bruto)
            t1 = time.time()
            if comprimido:
                bruto = gzip.decompress(bruto)
            corpo = bruto.decode("utf-8-sig")
            dados = json.loads(corpo) if corpo.strip() else []
            st["leitura"] += time.time() - t1
            if isinstance(dados, dict):
                raise RuntimeError(str(dados.get("error") or dados.get("erro") or dados.get("message") or dados))
            return dados
        except urllib.error.HTTPError as e:
            st["req"] += 1
            st["rede"] += time.time() - t0
            if e.code == 404:
                raise ErroHTTP(404, "série não encontrada (HTTP 404)")
            if e.code in (400, 406):            # consulta recusada: repetir não resolve
                raise ErroHTTP(e.code, "consulta recusada pela API (HTTP %s)" % e.code)
            ultimo = e
            if e.code == 429:                   # limite de requisições da API
                espera = _espera_429(e, t)
        except (socket.timeout, TimeoutError):  # conectou, mas o servidor não responde
            st["req"] += 1
            st["rede"] += time.time() - t0
            raise ErroSemResposta("sem resposta do servidor em %.0f s%s" % (limite, dica))
        except (OSError, ValueError) as e:      # falha de conexão, JSON inválido
            st["rede"] += time.time() - t0
            ultimo = e
        if t < tentativas - 1:
            st["extras"] += 1
            time.sleep(espera)
    raise RuntimeError("falha de comunicação com a API: %s" % ultimo)


def _extrair(dados, valores):
    for item in dados:
        bruto = item.get("valor")
        if bruto in (None, ""):
            continue
        d, m, y = item["data"].split("/")
        valores[_serial(date(int(y), int(m), int(d)))] = float(bruto)  # API usa ponto decimal


DICA_INEXISTENTE = " (a série existe? o servidor não responde a séries inexistentes)"
DICA_PESADA = " (consulta pesada ou servidor lento: aumente TIMEOUT_SEGUNDOS ou reduza JANELA_ANOS)"


def _verificar_existencia(sid, stats):
    """Pede só o último valor da série: série que existe responde na hora; a inexistente não responde.
    Só 'sem resposta' e 404 reprovam; qualquer outra falha deixa a requisição principal decidir."""
    try:
        _get_json(URL_ULTIMOS.format(sid=sid), tentativas=1, stats=stats,
                  timeout=TIMEOUT_EXISTENCIA_SEGUNDOS, dica=DICA_INEXISTENTE)
    except ErroSemResposta:
        raise
    except ErroHTTP as e:
        if e.codigo == 404:
            raise
    except RuntimeError:
        pass


def baixar_serie(sid, ini, fim, stats=None):
    """Devolve [("", {número_de_série_da_data: valor})].

    Sem datas: tenta a série inteira de uma vez. Se a API recusar, refaz em janelas de 10 anos
    do mais recente para o mais antigo e para quando uma janela vem vazia depois de já ter
    dados (o início da série). Série inexistente (404) falha na hora, sem fallback.
    """
    valores = {}
    if VERIFICAR_EXISTENCIA:
        _verificar_existencia(sid, stats)
    if ini is None and fim is None:
        try:
            _extrair(_get_json(URL_SEM_FILTRO.format(sid=sid), stats=stats, dica=DICA_PESADA), valores)
            return [("", valores)]
        except ErroHTTP as e:
            if e.codigo == 404:
                raise
            valores.clear()
        except ErroSemResposta:
            raise                    # sem resposta: as janelas também ficariam sem resposta
        except RuntimeError:
            valores.clear()          # recusado ou instável: refaz em janelas
    fim = fim or date.today()
    if ini is None:
        for a, b in _janelas_retroativas(fim, DATA_MINIMA):
            antes = len(valores)
            _extrair(_get_json(URL.format(sid=sid, ini=a.strftime("%d/%m/%Y"), fim=b.strftime("%d/%m/%Y")),
                               stats=stats, dica=DICA_PESADA), valores)
            if valores and len(valores) == antes:     # janela vazia depois de já ter dados: acabou o histórico
                break
    else:
        for a, b in _janelas(ini, fim):
            _extrair(_get_json(URL.format(sid=sid, ini=a.strftime("%d/%m/%Y"), fim=b.strftime("%d/%m/%Y")),
                               stats=stats, dica=DICA_PESADA), valores)
    return [("", valores)]


# ---------------------------------------------------------------- rede: IPv4 primeiro e conexão rápida (comum)

def _instalar_rede():
    """Dois ajustes de rede, aplicados só aos hosts das APIs (HOSTS_API):

    1) IPv4 primeiro (PREFERIR_IPV4): se o IPv6 estiver bloqueado, o Python esperava o timeout inteiro
       antes de cair para o IPv4.
    2) Conexão rápida: o Python tenta os endereços do servidor um de cada vez e espera o timeout INTEIRO
       em cada um. Se um dos endereços estiver fora do ar (comum quando o DNS alterna entre vários), cada
       requisição que cai nele trava. Aqui cada endereço recebe CONEXAO_RAPIDA_SEGUNDOS e o último que
       funcionou passa a ser tentado primeiro. Se TODOS falharem, repete com o timeout normal.
    """
    if not hasattr(socket, "_raf_original"):
        socket._raf_original = socket.getaddrinfo
        socket._raf_hosts = set()                      # hosts com IPv4 primeiro

        def getaddrinfo(host, *args, **kwargs):
            resultado = socket._raf_original(host, *args, **kwargs)
            if host in socket._raf_hosts:
                return sorted(resultado, key=lambda r: r[0] != socket.AF_INET)   # ordenação estável: IPv4 antes
            return resultado
        socket.getaddrinfo = getaddrinfo
    if not hasattr(socket, "_raf_create_original"):
        socket._raf_create_original = socket.create_connection
        socket._raf_hosts_rapidos = set()              # hosts com conexão rápida
        socket._raf_bons = {}                          # host -> endereço que conectou da última vez

        def create_connection(endereco, timeout=socket._GLOBAL_DEFAULT_TIMEOUT, source_address=None, *args, **kwargs):
            host, porta = endereco[0], endereco[1]
            if host not in socket._raf_hosts_rapidos:
                return socket._raf_create_original(endereco, timeout, source_address, *args, **kwargs)
            leitura = timeout if isinstance(timeout, (int, float)) else socket.getdefaulttimeout()
            infos = socket.getaddrinfo(host, porta, 0, socket.SOCK_STREAM)
            bom = socket._raf_bons.get(host)
            if bom:                                    # o que funcionou da última vez vai na frente
                infos = sorted(infos, key=lambda r: r[4][0] != bom)
            passadas = [CONEXAO_RAPIDA_SEGUNDOS if not leitura else min(CONEXAO_RAPIDA_SEGUNDOS, leitura)]
            if leitura and leitura > passadas[0]:
                passadas.append(leitura)               # 2ª passada, com a paciência normal, só se a 1ª falhar em todos
            ultimo = None
            for tempo in passadas:
                for familia, tipo, proto, _, sockaddr in infos:
                    s = socket.socket(familia, tipo, proto)
                    try:
                        s.settimeout(tempo)
                        if source_address:
                            s.bind(source_address)
                        s.connect(sockaddr)
                        s.settimeout(leitura)          # volta ao tempo de espera de leitura
                        socket._raf_bons[host] = sockaddr[0]
                        return s
                    except OSError as e:
                        ultimo = e
                        s.close()
            raise ultimo or OSError("nenhum endereço para %s" % host)
        socket.create_connection = create_connection

    if PREFERIR_IPV4:
        socket._raf_hosts.update(HOSTS_API)
    else:
        socket._raf_hosts.difference_update(HOSTS_API)
    socket._raf_hosts_rapidos.update(HOSTS_API)


_instalar_rede()


# ---------------------------------------------------------------- datas (comum)

def _data_de_serial(v):
    """Célula de data do Calc (número) ou texto DD/MM/AAAA -> date; vazio -> None."""
    if isinstance(v, (int, float)):
        return EPOCA_CALC + timedelta(days=int(v)) if v else None
    texto = str(v).strip()
    if not texto:
        return None
    try:
        return datetime.strptime(texto, "%d/%m/%Y").date()
    except ValueError:
        raise ValueError("data inválida '%s' (use DD/MM/AAAA)" % texto)


# ---------------------------------------------------------------- planilha (comum)

STATUS_COR = {"ok": 0xC6EFCE, "erro": 0xFFC7CE, "aviso": 0xFFEB9C}
_EM_EXECUCAO = False
_INICIO_EXECUCAO = 0.0
_REFS = []          # mantém vivas as referências usadas no modo em segundo plano

try:
    import unohelper
    from com.sun.star.awt import XCallback

    class _Retorno(unohelper.Base, XCallback):
        """Executa uma função na thread principal do LibreOffice (AsyncCallback)."""

        def __init__(self, funcao):
            self._funcao = funcao

        def notify(self, dados):
            try:
                self._funcao()
            finally:
                if self in _REFS:
                    _REFS.remove(self)
except ImportError:         # fora do LibreOffice (testes)
    _Retorno = None


def _obter_aba(doc, nome):
    if not doc.Sheets.hasByName(nome):
        doc.Sheets.insertNewByName(nome, len(doc.Sheets))
    return doc.Sheets.getByName(nome)


def _ler_parametros(aba):
    """Lê A2:D... de uma vez. Devolve (series, status); status = {linha: (texto, tipo)}."""
    faixa = aba.getCellRangeByPosition(0, 1, 3, MAX_LINHAS_PARAMETROS)
    series, status = [], {}
    for n, linha in enumerate(faixa.getDataArray(), start=2):
        if str(linha[0]).strip() == "":
            continue
        try:
            sid = _normalizar_id(linha[0])
            ini, fim = _data_de_serial(linha[2]), _data_de_serial(linha[3])
            if ini and fim and ini > fim:
                raise ValueError("data inicial maior que a final")
        except ValueError as e:
            status[n] = ("ERRO: %s" % e, "erro")
            continue
        desc = str(linha[1]).strip()
        series.append({"linha": n, "id": sid, "nome": desc or sid, "ini": ini, "fim": fim})
    return series, status


def _escrever_status(aba, status):
    """Grava a coluna E (Status) das linhas 2..N e pinta: verde = ok, amarelo = aviso, vermelho = erro."""
    faixa = aba.getCellRangeByPosition(4, 1, 4, MAX_LINHAS_PARAMETROS)
    faixa.setDataArray(tuple((status.get(n, ("", None))[0],)
                             for n in range(2, MAX_LINHAS_PARAMETROS + 2)))
    faixa.CellBackColor = -1                      # sem cor
    for n, (_, tipo) in status.items():
        if tipo in STATUS_COR:
            aba.getCellByPosition(4, n - 1).CellBackColor = STATUS_COR[tipo]
    aba.Columns.getByIndex(4).OptimalWidth = True


def _log(aba, msg):
    celula = aba.getCellByPosition(5, 1)    # F2
    celula.String = msg
    celula.CellBackColor = 0xD3D3D3
    celula.CharHeight = 12
    celula.HoriJustify = 1
    celula.VertJustify = 1
    aba.Columns.getByIndex(5).OptimalWidth = True
    aba.Rows.getByIndex(1).OptimalHeight = True


def _ler_tabela_anterior(aba):
    """Lê a tabela atual do destino. Devolve ({cabeçalho: {data: valor}}, dados brutos, área usada)."""
    cursor = aba.createCursor()
    cursor.gotoEndOfUsedArea(False)
    ult = cursor.RangeAddress
    dados = aba.getCellRangeByPosition(0, 0, max(ult.EndColumn, 0), max(ult.EndRow, 0)).getDataArray()
    antigo = {}
    if len(dados) > 1 and len(dados[0]) > 1:
        for j in range(1, len(dados[0])):
            coluna = {}
            for r in range(1, len(dados)):
                d, v = dados[r][0], dados[r][j]
                if isinstance(d, float) and isinstance(v, float):
                    coluna[d] = v
            antigo[str(dados[0][j])] = coluna
    return antigo, dados, ult


def _escrever(doc, aba, nomes, resultados, ult, backup):
    datas = sorted(set().union(*(r.keys() for r in resultados)))
    dados = [("Data",) + tuple(nomes)]
    for d in datas:
        dados.append((d,) + tuple(r.get(d, "") for r in resultados))   # "" = célula vazia
    n_lin, n_col = len(dados), len(nomes) + 1
    try:
        # tudo já está montado em memória; só agora a tabela antiga é apagada
        aba.getCellRangeByPosition(0, 0, max(ult.EndColumn, 0), max(ult.EndRow, 0)).clearContents(LIMPAR)
        aba.getCellRangeByPosition(0, 0, n_col - 1, n_lin - 1).setDataArray(tuple(dados))

        locale = uno.createUnoStruct("com.sun.star.lang.Locale")
        locale.Language, locale.Country = "pt", "BR"
        formatos = doc.NumberFormats
        k_data = formatos.getFormatIndex(
            uno.getConstantByName("com.sun.star.i18n.NumberFormatIndex.DATE_SYS_DDMMYYYY"), locale)
        k_num = formatos.getFormatIndex(
            uno.getConstantByName("com.sun.star.i18n.NumberFormatIndex.NUMBER_1000DEC2"), locale)

        cab = aba.getCellRangeByPosition(0, 0, n_col - 1, 0)
        cab.CellBackColor = 0xADD8E6
        cab.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
        if n_lin > 1:
            aba.getCellRangeByPosition(0, 1, 0, n_lin - 1).NumberFormat = k_data
            if n_col > 1:
                aba.getCellRangeByPosition(1, 1, n_col - 1, n_lin - 1).NumberFormat = k_num
        cab.Columns.OptimalWidth = True
    except Exception:
        try:    # tenta devolver o que havia antes
            aba.getCellRangeByPosition(0, 0, n_col - 1, n_lin - 1).clearContents(LIMPAR)
            aba.getCellRangeByPosition(0, 0, max(ult.EndColumn, 0), max(ult.EndRow, 0)).setDataArray(backup)
        except Exception:
            pass
        raise
    return n_lin - 1


# ---------------------------------------------------------------- execução (comum)

def _baixar_todas(series, progresso=None):
    """Baixa todas as séries em paralelo. Devolve [(partes, erro)] na ordem de 'series'."""
    def tarefa(s):
        stats = {"req": 0, "rede": 0.0, "leitura": 0.0, "bytes": 0, "extras": 0}
        try:
            return baixar_serie(s["id"], s["ini"], s["fim"], stats), None, stats
        except RuntimeError as e:
            return None, str(e), stats
        except Exception as e:      # erro inesperado de uma série não derruba as outras
            return None, "%s: %s" % (type(e).__name__, e), stats

    retornos = [None] * len(series)
    with ThreadPoolExecutor(max_workers=max(1, min(MAX_THREADS, len(series)))) as pool:
        futuros = {pool.submit(tarefa, s): i for i, s in enumerate(series)}
        for n, fut in enumerate(as_completed(futuros), start=1):
            retornos[futuros[fut]] = fut.result()
            if progresso:
                progresso.passo(n)
    return retornos


class _Progresso:
    """Barra de status do LibreOffice (só informativa)."""

    def __init__(self, doc, total):
        try:
            self.ind = doc.CurrentController.StatusIndicator
            self.ind.start("Baixando séries (%s)..." % NOME_FONTE, total)
        except Exception:
            self.ind = None

    def passo(self, n):
        try:
            if self.ind:
                self.ind.setValue(n)
        except Exception:
            pass

    def fim(self):
        try:
            if self.ind:
                self.ind.end()
        except Exception:
            pass


def _resumo_stats(stats):
    """Texto curto: nº de requisições, tempo de rede, tamanho e novas tentativas."""
    if not stats or not stats.get("req"):
        return ""
    txt = " [%d req, rede %.1f s, %.0f KB" % (stats["req"], stats["rede"], stats["bytes"] / 1024.0)
    if stats.get("extras"):
        txt += ", %d nova(s) tentativa(s)" % stats["extras"]
    return txt + "]"


def _aviso_final(doc, texto, gancho=None):
    """Avisa que a atualização em segundo plano terminou, sem interromper o trabalho.
    AVISO_FINAL = "BARRA": mensagem na barra de status por alguns segundos; "CAIXA": janela com OK; None: só o log em F2."""
    if not AVISO_FINAL:
        return False
    try:
        if AVISO_FINAL == "CAIXA":
            ctx = uno.getComponentContext()
            tk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
            tk.createMessageBox(doc.CurrentController.Frame.ContainerWindow,
                                uno.Enum("com.sun.star.awt.MessageBoxType", "INFOBOX"), 1, NOME_FONTE, texto).execute()
            return True
        indicador = doc.CurrentController.StatusIndicator
        indicador.start(texto, 1)
        indicador.setValue(1)
        if gancho is not None and _Retorno is not None:      # some sozinho, na thread principal, depois de AVISO_SEGUNDOS
            def encerrar():
                try:
                    indicador.end()
                except Exception:
                    pass

            def esperar():
                time.sleep(AVISO_SEGUNDOS)
                objeto = _Retorno(encerrar)
                _REFS.append(objeto)
                try:
                    gancho.addCallback(objeto, None)
                except Exception:
                    pass
            threading.Thread(target=esperar, daemon=True).start()
        else:
            indicador.end()
        return True
    except Exception:
        return False


def _concluir(doc, aba_par, aba_dest, series, status, retornos, t_download=None):
    """Monta a tabela e os status. Roda sempre na thread principal."""
    agora = datetime.now().strftime("%H:%M:%S")
    anterior, backup, ult = _ler_tabela_anterior(aba_dest)
    nomes, resultados = [], []
    n_ok = 0
    for s, (partes, erro, stats) in zip(series, retornos):
        if erro is None:
            n_ok += 1
            linhas = len(set().union(*(v.keys() for _, v in partes))) if partes else 0
            for terr, valores in partes:
                nomes.append(s["nome"] if not terr else "%s [%s]" % (s["nome"], terr))
                resultados.append(valores)
            if linhas == 0:
                status[s["linha"]] = ("AVISO: sem dados no período (%s)" % agora, "aviso")
            else:
                status[s["linha"]] = ("OK: %d coluna(s), %d linha(s) (%s)%s" % (
                    len(partes), linhas, agora, _resumo_stats(stats)), "ok")
        else:
            antigas = [(n, c) for n, c in anterior.items()
                       if n == s["nome"] or n.startswith(s["nome"] + " [")]
            for n, c in antigas:        # mantém a coluna que já existia na tabela
                nomes.append(n)
                resultados.append(c)
            status[s["linha"]] = ("ERRO: %s | %s (%s)%s" % (
                erro, "dados anteriores mantidos" if antigas else "sem dados anteriores", agora,
                _resumo_stats(stats)), "erro")

    n_erro = len(series) - n_ok
    if n_ok == 0:
        _escrever_status(aba_par, status)
        _log(aba_par, "Nenhuma série baixada; a tabela '%s' não foi alterada (%s)." % (ABA_DESTINO, agora))
        return

    doc.lockControllers()
    doc.enableAutomaticCalculation(False)
    t_gravacao = time.time()
    t_pos = 0.0
    resumo = ""
    try:
        n = _escrever(doc, aba_dest, nomes, resultados, ult, backup)
        t_gravacao = time.time() - t_gravacao
        resumo = "Atualizado em %s: %d série(s) ok, %d com erro, %d linha(s). Tempo: download %s, gravação %.1f s." % (
            datetime.now().strftime("%d/%m/%Y %H:%M:%S"), n_ok, n_erro, n,
            ("%.1f s" % t_download) if t_download is not None else "?", t_gravacao)
    except Exception as e:
        for s in series:
            if status[s["linha"]][1] in ("ok", "aviso"):
                status[s["linha"]] = ("ERRO: falha ao gravar a tabela (%s)" % e, "erro")
        resumo = "Falha ao gravar a tabela; o conteúdo anterior foi restaurado."
    finally:
        t_pos = time.time()
        doc.enableAutomaticCalculation(True)         # aqui o LibreOffice recalcula as fórmulas que dependem da tabela
        doc.unlockControllers()
        t_pos = time.time() - t_pos
    resumo += " Recálculo e redesenho: %.1f s." % t_pos
    _escrever_status(aba_par, status)
    _log(aba_par, resumo)


def _executar():
    global _EM_EXECUCAO, _INICIO_EXECUCAO
    if _EM_EXECUCAO and time.time() - _INICIO_EXECUCAO < LIMITE_EXECUCAO_SEGUNDOS:
        return                      # já há uma atualização em andamento
    _EM_EXECUCAO = True             # (se a anterior passou do limite sem terminar, considera-se travada e começa outra)
    _INICIO_EXECUCAO = inicio_macro = time.time()
    assincrono = False
    aba_par = None
    try:
        ctx = uno.getComponentContext()
        doc = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx).getCurrentComponent()
        aba_par = _obter_aba(doc, ABA_PARAMETROS)
        aba_dest = _obter_aba(doc, ABA_DESTINO)

        cab = aba_par.getCellRangeByPosition(0, 0, 5, 0)
        cab.setDataArray(((ROTULO_ID, "Descrição (opcional)", "Data Inicial (DD/MM/AAAA)",
                           "Data Final (DD/MM/AAAA)", "Status", "Log"),))
        cab.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")

        series, status = _ler_parametros(aba_par)
        if not series:
            _escrever_status(aba_par, status)
            _log(aba_par, "Nenhuma série válida em %s (coluna A, a partir da linha 2)." % ABA_PARAMETROS)
            return
        for s in series:
            status[s["linha"]] = ("Baixando...", None)
        _escrever_status(aba_par, status)

        tempos = {"inicio": inicio_macro}

        def concluir(retornos):
            global _EM_EXECUCAO
            try:
                if isinstance(retornos, Exception):
                    retornos = [(None, "falha geral: %s" % retornos, {})] * len(series)
                _concluir(doc, aba_par, aba_dest, series, status, retornos, tempos.get("download"))
                try:
                    celula = aba_par.getCellByPosition(5, 1)
                    celula.String = celula.String + " Total da macro: %.1f s." % (time.time() - tempos["inicio"])
                except Exception:
                    pass
                if tempos.get("gancho") is not None:               # só no modo em segundo plano: o usuário não estava esperando
                    n_ok = sum(1 for v in status.values() if v[1] == "ok")
                    n_av = sum(1 for v in status.values() if v[1] == "aviso")
                    n_er = sum(1 for v in status.values() if v[1] == "erro")
                    _aviso_final(doc, "%s: atualização concluída - %d ok, %d aviso(s), %d com erro"
                                 % (NOME_FONTE, n_ok, n_av, n_er), tempos["gancho"])
            except Exception as e:
                try:
                    _log(aba_par, "Erro inesperado: %s: %s" % (type(e).__name__, e))
                except Exception:
                    pass
            finally:
                _EM_EXECUCAO = False

        if EXECUCAO_ASSINCRONA and _Retorno is not None:
            gancho = ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.AsyncCallback", ctx)
            tempos["gancho"] = gancho

            def trabalho():          # só rede e Python puro: nada de UNO aqui, exceto o agendamento final
                global _EM_EXECUCAO
                t_ini = time.time()
                try:
                    retornos = _baixar_todas(series)
                except Exception as e:
                    retornos = e
                tempos["download"] = time.time() - t_ini
                obj = _Retorno(lambda: concluir(retornos))
                _REFS.append(obj)
                try:
                    gancho.addCallback(obj, None)
                except Exception:
                    try:
                        concluir(retornos)            # último recurso: grava daqui mesmo, em vez de perder o resultado
                    except Exception:
                        pass
                    _EM_EXECUCAO = False

            _log(aba_par, "Baixando em segundo plano... (%s)" % datetime.now().strftime("%H:%M:%S"))
            threading.Thread(target=trabalho, daemon=True).start()
            assincrono = True
        else:
            progresso = _Progresso(doc, len(series))
            t_ini = time.time()
            try:
                retornos = _baixar_todas(series, progresso)
            finally:
                progresso.fim()
            tempos["download"] = time.time() - t_ini
            concluir(retornos)
    except Exception as e:
        if aba_par is not None:
            try:
                _log(aba_par, "Erro inesperado: %s: %s" % (type(e).__name__, e))
            except Exception:
                pass
    finally:
        if not assincrono:
            _EM_EXECUCAO = False


# ---------------------------------------------------------------- macro

def AtualizarSGS(*args):
    """Atualiza as séries listadas em 'Parametros_SGS' na aba 'Tabela_SGS'."""
    _executar()


g_exportedScripts = (AtualizarSGS,)
