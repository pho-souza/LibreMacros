# -*- coding: utf-8 -*-
"""Diagnóstico de rede para as macros SGS e Ipeadata no LibreOffice.

Rode DiagnosticoRede (Ferramentas > Macros > Executar macro > Minhas macros > diagnostico_rede).
Ela mede, na SUA máquina e com o Python do LibreOffice, onde o tempo de uma requisição vai:

  DNS ........ resolver o nome do servidor e quais endereços (IPv4/IPv6) ele devolve
  Conexão .... tempo para conectar em CADA endereço (um endereço fora do ar ou bloqueado estoura aqui)
  Sondagem ... 6 conexões seguidas, cada uma com DNS novo: se o DNS alterna entre endereços e um deles
               está morto, só parte das conexões demora (é o que faz a macro "às vezes" travar)
  Requisição . conexão (DNS + TCP + TLS), espera até o 1º byte, download e leitura do JSON,
               duas vezes: na ordem de endereços do sistema e com IPv4 primeiro
  Correção ... 3 requisições seguidas com a conexão rápida das macros (cada endereço recebe 2 s e o que
               funcionou é lembrado): mostra se o travamento acabou

O resultado aparece na aba 'Diagnostico_rede' do arquivo aberto (dá para copiar e colar aqui).
Não altera nenhuma outra aba. Usa só leitura (GET) nas séries 189 do SGS e PRECOS12_IPCA12 do Ipeadata.
NÃO foi testado contra as APIs reais (o ambiente onde foi escrito não alcança esses servidores).
"""
import gzip
import http.client
import json
import socket
import time
import urllib.parse
import urllib.request

try:
    import uno
except ImportError:             # permite testar a lógica fora do LibreOffice
    uno = None

ALVOS = [
    ("BCB / SGS - IGP-M (série 189), sem datas",
     "https://api.bcb.gov.br/dados/serie/bcdata.sgs.189/dados?formato=json"),
    ("Ipeadata - IPCA (PRECOS12_IPCA12)",
     "https://www.ipeadata.gov.br/api/odata4/ValoresSerie(SERCODIGO='PRECOS12_IPCA12')"),
]
TIMEOUT = 20.0          # espera máxima por resposta
TIMEOUT_CONEXAO = 5.0   # espera máxima ao testar cada endereço
SONDAGENS = 6           # nº de conexões seguidas na sondagem
CONEXAO_RAPIDA = 2.0    # espera por cada endereço na conexão rápida (igual às macros)
TIMEOUT_SONDA = 4.0     # espera máxima de cada conexão da sondagem
LENTO = 1.5             # acima disso (segundos) uma etapa é considerada lenta


def _familia(f):
    return "IPv4" if f == socket.AF_INET else ("IPv6" if f == socket.AF_INET6 else str(f))


def _resolvedor():
    """Resolvedor original, mesmo que as macros SGS/Ipeadata tenham instalado a ordenação IPv4."""
    return getattr(socket, "_raf_original", socket.getaddrinfo)


def medir_dns_e_conexoes(host, porta):
    """Devolve (linhas, tempo_dns, [(familia, ip, tempo, ok)])."""
    linhas, conexoes = [], []
    t = time.time()
    try:
        infos = _resolvedor()(host, porta, 0, socket.SOCK_STREAM)
    except OSError as e:
        return ["DNS: FALHOU (%s)" % e], None, conexoes
    t_dns = time.time() - t
    vistos = []
    for fam, _, _, _, addr in infos:
        if (fam, addr[0]) not in vistos:
            vistos.append((fam, addr[0]))
    linhas.append("DNS: %.2f s, %d endereço(s), ordem devolvida pelo sistema: %s" % (
        t_dns, len(vistos), ", ".join("%s %s" % (_familia(f), ip) for f, ip in vistos)))
    for fam, ip in vistos[:6]:
        s = socket.socket(fam, socket.SOCK_STREAM)
        s.settimeout(TIMEOUT_CONEXAO)
        t = time.time()
        try:
            s.connect((ip, porta))
            dt, ok, txt = time.time() - t, True, "ok"
        except OSError as e:
            dt, ok, txt = time.time() - t, False, "FALHOU (%s)" % type(e).__name__
        finally:
            s.close()
        conexoes.append((fam, ip, dt, ok))
        linhas.append("  conectar em %s %s: %s em %.2f s" % (_familia(fam), ip, txt, dt))
    return linhas, t_dns, conexoes


def requisicao(url, ipv4_primeiro, manter_conexao=False):
    """Uma requisição GET medida por etapa. Devolve dicionário com tempos (s), bytes e erro."""
    partes = urllib.parse.urlparse(url)
    host, https = partes.hostname, partes.scheme == "https"
    porta = partes.port or (443 if https else 80)
    caminho = partes.path + (("?" + partes.query) if partes.query else "")

    base = socket.getaddrinfo                       # troca só durante esta medição
    base_conexao = socket.create_connection         # mede o comportamento ORIGINAL do Python, sem as correções das macros
    if not manter_conexao:
        socket.create_connection = getattr(socket, "_raf_create_original", base_conexao)
    resolvedor = _resolvedor()

    def getaddrinfo(h, *a, **k):
        res = resolvedor(h, *a, **k)
        return sorted(res, key=lambda r: r[0] != socket.AF_INET) if (ipv4_primeiro and h == host) else res

    r = {"erro": None}
    socket.getaddrinfo = getaddrinfo
    try:
        cls = http.client.HTTPSConnection if https else http.client.HTTPConnection
        conn = cls(host, porta, timeout=TIMEOUT)
        t0 = time.time()
        conn.connect()
        r["conexao"] = time.time() - t0              # DNS + TCP + TLS
        t1 = time.time()
        conn.request("GET", caminho, headers={"User-Agent": "Mozilla/5.0 (diagnostico LibreOffice)",
                                              "Accept": "application/json", "Accept-Encoding": "gzip"})
        resp = conn.getresponse()
        r["status"] = resp.status
        r["espera"] = time.time() - t1               # até o 1º byte da resposta
        t2 = time.time()
        corpo = resp.read()
        r["download"] = time.time() - t2
        r["bytes"] = len(corpo)
        r["gzip"] = (resp.getheader("Content-Encoding") or "").lower() == "gzip"
        t3 = time.time()
        if r["gzip"]:
            corpo = gzip.decompress(corpo)
        try:
            dados = json.loads(corpo.decode("utf-8-sig"))
            r["registros"] = len(dados.get("value", dados)) if isinstance(dados, dict) else len(dados)
        except ValueError:
            r["registros"] = None
        r["leitura"] = time.time() - t3
        conn.close()
    except Exception as e:                           # qualquer falha de rede vira resultado, não exceção
        r["erro"] = "%s: %s" % (type(e).__name__, e)
    finally:
        socket.getaddrinfo = base
        socket.create_connection = base_conexao
    r["total"] = sum(r.get(k, 0.0) for k in ("conexao", "espera", "download", "leitura"))
    return r


def sondar(host, porta):
    """N conexões TCP seguidas, cada uma com DNS novo, no comportamento original do Python.
    Devolve (linhas, [(ip, tempo, ok)])."""
    criar = getattr(socket, "_raf_create_original", socket.create_connection)
    resolver = _resolvedor()
    base = socket.getaddrinfo
    resultados = []
    socket.getaddrinfo = lambda h, *a, **k: resolver(h, *a, **k)
    try:
        for _ in range(SONDAGENS):
            t = time.time()
            ip, ok = "?", False
            try:
                s = criar((host, porta), TIMEOUT_SONDA)
                ip, ok = s.getpeername()[0], True
                s.close()
            except OSError as e:
                ip = type(e).__name__
            resultados.append((ip, time.time() - t, ok))
    finally:
        socket.getaddrinfo = base
    lentas = [r for r in resultados if (not r[2]) or r[1] > LENTO]
    maximo = max(r[1] for r in resultados)
    linhas = ["Sondagem: %d conexões seguidas -> %d lenta(s) ou falha(s), a pior com %.2f s (tempos: %s)" % (
        SONDAGENS, len(lentas), maximo, ", ".join("%.2f" % r[1] for r in resultados))]
    return linhas, resultados


def _fabrica_conexao_rapida():
    """Réplica da conexão rápida das macros SGS/Ipeadata, com memória própria (começa vazia)."""
    bons = {}
    resolver = _resolvedor()

    def criar(endereco, timeout=TIMEOUT, source_address=None, *a, **k):
        host, porta = endereco[0], endereco[1]
        infos = sorted(resolver(host, porta, 0, socket.SOCK_STREAM), key=lambda r: r[0] != socket.AF_INET)   # IPv4 primeiro
        if bons.get(host):
            infos = sorted(infos, key=lambda r: r[4][0] != bons[host])                                       # o que funcionou, na frente
        ultimo = None
        for tempo in (CONEXAO_RAPIDA, timeout):
            for fam, tipo, proto, _, sa in infos:
                sk = socket.socket(fam, tipo, proto)
                try:
                    sk.settimeout(tempo)
                    sk.connect(sa)
                    sk.settimeout(timeout)
                    bons[host] = sa[0]
                    return sk
                except OSError as e:
                    ultimo, _ = e, sk.close()
        raise ultimo or OSError("nenhum endereço")
    return criar


def requisicoes_com_correcao(url, n=3):
    """n requisições seguidas com a conexão rápida (memória vazia no começo). Devolve lista de resultados."""
    criar = _fabrica_conexao_rapida()
    base = socket.create_connection
    socket.create_connection = criar
    try:
        return [requisicao(url, ipv4_primeiro=True, manter_conexao=True) for _ in range(n)]
    finally:
        socket.create_connection = base


def _linha_requisicao(rotulo, r):
    if r["erro"]:
        return "%s: FALHOU - %s" % (rotulo, r["erro"])
    return ("%s: total %.2f s = conexão %.2f + espera do servidor %.2f + download %.2f + leitura %.2f | "
            "HTTP %s, %.0f KB%s, %s registro(s)" % (
                rotulo, r["total"], r["conexao"], r["espera"], r["download"], r["leitura"], r.get("status"),
                r["bytes"] / 1024.0, " (gzip)" if r["gzip"] else " (sem gzip)", r.get("registros")))


def conclusao(r_sistema, r_ipv4, conexoes, sondagem=None):
    """Texto curto apontando a etapa mais lenta."""
    msgs = []
    mortos = [c for c in conexoes if not c[3]]
    vivos = [c for c in conexoes if c[3] and c[2] <= LENTO]
    if mortos and vivos and not any(c[0] == socket.AF_INET6 for c in mortos):
        msgs.append("O endereço %s NÃO conecta (timeout) enquanto %s conecta. O DNS alterna a ordem desses endereços, então parte das "
                    "conexões cai no endereço morto e espera o timeout inteiro: é isso que faz a macro travar 'às vezes'. O navegador "
                    "tenta os dois quase ao mesmo tempo e não sente. As macros agora tentam cada endereço com timeout curto e lembram "
                    "o que funciona (CONEXAO_RAPIDA_SEGUNDOS)." % (", ".join(c[1] for c in mortos), ", ".join(c[1] for c in vivos)))
    if sondagem:
        lentas = [r for r in sondagem if (not r[2]) or r[1] > LENTO]
        if lentas and len(lentas) < len(sondagem):
            msgs.append("Na sondagem, %d de %d conexões foram lentas: problema intermitente (típico de endereço morto no rodízio do DNS)."
                        % (len(lentas), len(sondagem)))
    lentas = [c for c in conexoes if (not c[3]) or c[2] > LENTO]
    if any(c[0] == socket.AF_INET6 for c in lentas) and any(c[0] == socket.AF_INET and c[3] and c[2] <= LENTO for c in conexoes):
        msgs.append("O IPv6 deste servidor não conecta (ou é lento) e o IPv4 é rápido. É a causa clássica de 'Python lento, navegador "
                    "instantâneo'. As macros já tentam o IPv4 primeiro (PREFERIR_IPV4 = True).")
    ok = [r for r in (r_sistema, r_ipv4) if not r["erro"]]
    if not ok:
        msgs.append("A requisição falhou nas duas formas; veja o erro acima (proxy, certificado ou bloqueio de rede).")
        return msgs
    pior = max(ok, key=lambda r: r["total"])
    etapas = {"conexao": "a conexão (DNS + TCP + TLS: proxy, IPv6 ou certificados)",
              "espera": "a espera pela resposta do servidor",
              "download": "o download (tamanho da resposta / velocidade)",
              "leitura": "a leitura do JSON (resposta muito grande)"}
    etapa = max(etapas, key=lambda k: pior.get(k, 0.0))
    if pior["total"] > LENTO:
        msgs.append("Etapa mais lenta: %s (%.2f s de %.2f s)." % (etapas[etapa], pior[etapa], pior["total"]))
    else:
        msgs.append("A requisição é rápida aqui (%.2f s). Se a macro ainda demora, o gargalo não é a rede de uma "
                    "requisição: confira na coluna E de cada aba 'Parametros_*' o nº de requisições e as novas tentativas." % pior["total"])
    if r_sistema.get("total") and r_ipv4.get("total") and not r_sistema["erro"] and not r_ipv4["erro"]:
        if r_sistema["total"] - r_ipv4["total"] > LENTO:
            msgs.append("Com IPv4 primeiro a requisição é %.1f s mais rápida." % (r_sistema["total"] - r_ipv4["total"]))
    return msgs


def diagnosticar(alvos=None):
    """Roda o diagnóstico completo e devolve a lista de linhas do relatório."""
    linhas = ["Diagnóstico de rede - %s" % time.strftime("%d/%m/%Y %H:%M:%S"),
              "Python %s | proxies detectados: %s" % (
                  __import__("sys").version.split()[0], urllib.request.getproxies() or "nenhum")]
    for rotulo, url in (alvos or ALVOS):
        partes = urllib.parse.urlparse(url)
        porta = partes.port or (443 if partes.scheme == "https" else 80)
        linhas += ["", "== %s" % rotulo, "URL: %s" % url]
        l_dns, _, conexoes = medir_dns_e_conexoes(partes.hostname, porta)
        linhas += l_dns
        r_sis = requisicao(url, ipv4_primeiro=False)
        r_v4 = requisicao(url, ipv4_primeiro=True)
        linhas.append(_linha_requisicao("Requisição na ordem do sistema", r_sis))
        linhas.append(_linha_requisicao("Requisição com IPv4 primeiro  ", r_v4))
        l_sonda, sondagem = sondar(partes.hostname, porta)
        linhas += l_sonda
        com = requisicoes_com_correcao(url)
        if any(r["erro"] for r in com):
            linhas.append("Com a correção das macros: FALHOU - %s" % next(r["erro"] for r in com if r["erro"]))
        else:
            linhas.append("Com a correção das macros: 3 requisições seguidas em %s (a 1ª pode custar até %.0f s se o endereço morto vier primeiro; depois o endereço bom é lembrado)"
                          % (", ".join("%.2f s" % r["total"] for r in com), CONEXAO_RAPIDA))
        linhas += ["=> " + m for m in conclusao(r_sis, r_v4, conexoes, sondagem)]
        if not any(r["erro"] for r in com):
            pior = max(r["total"] for r in com)
            if pior <= CONEXAO_RAPIDA + 1.0 and max(r["total"] for r in (r_sis, r_v4) if not r["erro"]) > LENTO:
                linhas.append("=> A correção resolve: a pior das 3 requisições levou %.2f s (sem ela chegou a %.1f s)." % (
                    pior, max(r["total"] for r in (r_sis, r_v4) if not r["erro"])))
    return linhas


def DiagnosticoRede(*args):
    """Mede a rede e escreve o relatório na aba 'Diagnostico_rede' do arquivo aberto."""
    ctx = uno.getComponentContext()
    doc = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx).getCurrentComponent()
    linhas = diagnosticar()
    nome = "Diagnostico_rede"
    if not doc.Sheets.hasByName(nome):
        doc.Sheets.insertNewByName(nome, len(doc.Sheets))
    aba = doc.Sheets.getByName(nome)
    aba.getCellRangeByPosition(0, 0, 0, max(len(linhas), 60)).clearContents(1 | 2 | 4 | 16 | 32)
    aba.getCellRangeByPosition(0, 0, 0, len(linhas) - 1).setDataArray(tuple((l,) for l in linhas))
    aba.Columns.getByIndex(0).Width = 30000
    try:
        tk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
        tk.createMessageBox(doc.CurrentController.Frame.ContainerWindow,
                            uno.Enum("com.sun.star.awt.MessageBoxType", "INFOBOX"), 1, "Diagnóstico de rede",
                            "Pronto. O relatório está na aba 'Diagnostico_rede'. Copie e cole na conversa.").execute()
    except Exception:
        pass


g_exportedScripts = (DiagnosticoRede,)
