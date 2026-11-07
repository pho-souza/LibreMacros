# -*- coding: utf-8 -*-
"""Formatação de gráficos no LibreOffice Calc (port das macros VBA do suplemento RAF).

Macros expostas (Ferramentas > Macros > Executar macro > Minhas macros > formata_grafico):

  GraficosFonte           <- Graficos_Fonte (Ctrl+Alt+I): só formata
  GraficoTitulo           <- grafico_title (Ctrl+Alt+J): lê A3 (título), A4 (subtítulo), A5 (fonte),
                             liga os nomes das séries à linha 7 e depois formata. Célula em branco
                             é ignorada (o que já está no gráfico fica como está).
  GraficoTituloEnglish    <- grafico_title_english: título em A4, séries na linha 6, "Source"
  RotulosPorMes           <- AddSelectiveDataLabels (chart_format): rótulos só em um mês
  FormatarTodosGraficos   <- (novo) aplica GraficosFonte a todos os gráficos do arquivo
  TituloETodosGraficos    <- (novo) aplica GraficoTitulo a todos os gráficos do arquivo

COMO ESCOLHE O GRÁFICO
  1. Se há gráfico(s) selecionado(s) (um clique no gráfico), usa o(s) selecionado(s).
  2. Senão, se a aba ativa tem exatamente UM gráfico, usa esse (é o caso das abas 'Fig NN').
  3. Senão, pede para selecionar um gráfico. Nunca altera gráfico sem uma dessas condições.

O QUE A FORMATAÇÃO FAZ (valores nas constantes do bloco CONFIGURAÇÃO, logo abaixo)
  tamanho mínimo 18 x 10 cm (nunca diminui) | Calibri 10 preto | título Calibri 14 negrito em
  caixa-alta, no topo e centralizado | subtítulo (se houver) Calibri 10 colado ao título |
  fundo do gráfico #F5F5F5 | área de plotagem #EEEEEE sem borda | grade principal do eixo Y em
  #757575 com 0,01 cm (a do eixo X só é formatada se já existir) | eixo X cruzando no início, EXCETO
  nos gráficos com valores negativos, onde ele cruza em y = 0 (vira a linha preta) e os rótulos
  continuam embaixo | marcas principais externas, no eixo (não nos rótulos) | legenda embaixo, com 100% da largura (só se o gráfico já tiver legenda) | caixa "Fonte"
  no rodapé, criada com "Fonte: " se não existir.

LAYOUT DINÂMICO (centésimos de mm, calculado a cada execução)
  topo:   título -> subtítulo (se houver) -> área de plotagem; sem subtítulo a área sobe.
  rodapé: área de plotagem -> legenda -> caixa "Fonte"; a área termina acima da legenda, sem
          sobrepor os rótulos dos eixos. À direita sobra mais espaço quando o último ponto tem
          rótulo de dados.

PROTEÇÕES
  - uma célula de título que contenha número/data (ex.: 'fev/12') NUNCA vira título: vira aviso;
  - um cabeçalho de série (linha 7) que não seja texto não renomeia a série: vira aviso;
  - cada etapa é isolada; o que falhar aparece no aviso final, e a mensagem final lista o texto
    de título/subtítulo/fonte que foi lido de cada aba.

LINHA EM y = 0 (LINHA_ZERO, LINHA_ZERO_MODO)
  "EIXO" (padrão): o eixo X cruza em 0 com rótulos embaixo; é o desenho nativo do Excel
  (crosses=autoZero + tickLblPos=low) e sobrevive ao salvar em .xlsx.
  "SECUNDARIO": mantém o eixo X embaixo e desenha a linha com um eixo X secundário; funciona no
  LibreOffice/ODS, mas o LibreOffice DESCARTA esse eixo ao salvar em .xlsx.
  Só se aplica quando o zero está dentro da escala do eixo Y (há valores positivos e negativos).

OBSERVAÇÕES
  - no LibreOffice, definir uma célula como "" (String) não a esvazia; para limpar use Delete;
  - A4 e A5 em branco = ignora: um subtítulo ou fonte que já exista no gráfico NÃO é removido;
  - A4/A5 são lidos por GraficoTitulo (Ctrl+Alt+J) e pelas macros "Todos"; GraficosFonte
    (Ctrl+Alt+I) só formata e não lê células;
  - texto de A5 que não comece com "Fonte"/"Source" ganha "Fonte: " na frente (PREFIXAR_FONTE);
  - caixas de anotação (CaixaDeTexto) recebem Calibri/preto mas mantêm o tamanho de letra
    (ANOTACOES_TAMANHO_PROPRIO); no VBA viravam 10 e estouravam a caixa;
  - Fig 02 tem dois gráficos e um título só (A3): selecione cada gráfico e cuide do título à mão.

Testado em LibreOffice 24.2 (modo headless) com o arquivo RAF113_JUN2026; a seleção de gráfico
por clique (item 1 acima) só pôde ser simulada, não testada com mouse.
"""
import re
import time
from datetime import date, datetime, timedelta

import uno

# ======== CONFIGURAÇÃO (edite aqui) ========
PRETO = 0x000000
FONTE = "Calibri"
TAM_TEXTO = 10.0
TAM_TITULO = 14.0
TAM_SUBTITULO = 10.0
COR_SUBTITULO = PRETO
COR_FUNDO = 0xF5F5F5          # fundo do gráfico (área do gráfico)
PAREDE_SEM_BORDA = True       # remove a borda da parede (área de plotagem)
PAREDE_COR = 0xEEEEEE         # fundo da área de plotagem; None = não mexe
COR_GRADE = 0x757575          # grade principal
LINHA_ZERO = True             # linha preta em y = 0 (só nos gráficos cuja escala inclui o zero)
LINHA_ZERO_MODO = "EIXO"      # "EIXO": o eixo X cruza em 0 e os rótulos ficam embaixo (sobrevive ao .xlsx)
                              # "SECUNDARIO": eixo X embaixo + eixo secundário em 0 (SOME ao salvar em .xlsx)
COR_LINHA_ZERO = PRETO
ESPESSURA_LINHA_ZERO = 35     # 0,035 cm = 0,35 mm (~1 pt), em centésimos de mm
ESPESSURA_GRADE = 10          # 0,01 cm = 0,1 mm = 10 centésimos de mm (UNO usa 1/100 mm)
POSICAO_LEGENDA = "BOTTOM"    # LEFT, TOP, RIGHT ou BOTTOM
LEGENDA_LARGURA_TOTAL = True  # legenda com 100% da largura do gráfico (só com POSICAO_LEGENDA = "BOTTOM")
TEXTO_FONTE_PADRAO = "Fonte: "
CONGELAR_ESTILOS = True       # congela o redesenho do gráfico durante as etapas de estilo (mais rápido); o layout é calculado descongelado
PREFIXAR_FONTE = True         # texto de A5 sem "Fonte"/"Source" no início ganha "Fonte: " na frente
# células da planilha (GraficoTitulo): célula em branco = ignorada
CELULA_TITULO = "A3"
CELULA_SUBTITULO = "A4"
CELULA_FONTE = "A5"
LINHA_CABECALHO = 7           # linha com os nomes das séries
# dimensões, em centésimos de milímetro (unidade do UNO)
LARGURA_MIN = 18000           # 18 cm
ALTURA_MIN = 10000            # 10 cm
ALTURA_FONTE = 750            # 0,75 cm: altura da caixa "Fonte" colada no rodapé
FOLGA_SUBTITULO = 0           # espaço entre título e subtítulo (menor = mais junto)
FOLGA = 100                   # 1 mm entre legenda, caixa "Fonte" e área de plotagem
MARGEM_DIREITA = 300          # 0,3 cm livres à direita da área de plotagem
MARGEM_DIREITA_ROTULOS = 1500 # 1,5 cm quando o último ponto de alguma série tem rótulo de dados
ANOTACOES_TAMANHO_PROPRIO = True   # True: caixas de anotação (CaixaDeTexto) mantêm o tamanho de letra; False: viram Calibri 10
# ============================================
CLSID_GRAFICO = "12dcae26-281f-416f-a234-c3086127382e"
PREFIXOS_FONTE = ("fonte", "source")
NOME_CAIXA_FONTE = "LegendaGrafico"
EPOCA = date(1899, 12, 30)
SERIAL_MINIMO_DATA = 3660    # ~ 01/01/1910; abaixo disso o número é tratado como ano


class ErroUsuario(Exception):
    """Problema que deve ser explicado ao usuário (não é bug)."""


# ---------------------------------------------------------------- UNO: contexto e mensagens

def _contexto():
    ctx = uno.getComponentContext()
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return ctx, desktop.getCurrentComponent()


def _mensagem(doc, titulo, texto):
    try:
        ctx = uno.getComponentContext()
        tk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
        janela = doc.CurrentController.Frame.ContainerWindow
        caixa = tk.createMessageBox(janela, uno.Enum("com.sun.star.awt.MessageBoxType", "INFOBOX"),
                                    1, titulo, texto)
        caixa.execute()
    except Exception:
        pass


def _caixa_entrada(titulo, mensagem, padrao=""):
    """InputBox simples. Devolve o texto ou None se cancelar."""
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    modelo = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    modelo.Width, modelo.Height, modelo.Title = 150, 62, titulo

    def novo(tipo, nome, **props):
        m = modelo.createInstance("com.sun.star.awt.UnoControl%sModel" % tipo)
        for k, v in props.items():
            setattr(m, k, v)
        modelo.insertByName(nome, m)

    novo("FixedText", "rotulo", PositionX=6, PositionY=6, Width=138, Height=22, Label=mensagem, MultiLine=True)
    novo("Edit", "texto", PositionX=6, PositionY=30, Width=138, Height=12, Text=padrao)
    novo("Button", "ok", PositionX=36, PositionY=46, Width=36, Height=12, Label="OK",
         PushButtonType=1, DefaultButton=True)
    novo("Button", "cancelar", PositionX=78, PositionY=46, Width=36, Height=12, Label="Cancelar",
         PushButtonType=2)
    dialogo = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialogo.setModel(modelo)
    dialogo.createPeer(smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx), None)
    try:
        if dialogo.execute() == 1:
            return dialogo.getControl("texto").Text
        return None
    finally:
        dialogo.dispose()


# ---------------------------------------------------------------- localizar o gráfico

def _eh_grafico(forma):
    try:
        return (forma.supportsService("com.sun.star.drawing.OLE2Shape")
                and CLSID_GRAFICO in str(forma.CLSID).lower())
    except Exception:
        return False


def _formas_selecionadas(doc):
    sel = doc.CurrentController.getSelection()
    if sel is None:
        return []
    if _eh_grafico(sel):                       # uma forma só
        return [sel]
    achadas = []
    try:
        for i in range(sel.getCount()):        # coleção de formas (ou intervalos, que são ignorados)
            e = sel.getByIndex(i)
            if _eh_grafico(e):
                achadas.append(e)
    except Exception:
        pass
    return achadas


def _indices_de_graficos(aba):
    pagina = aba.DrawPage
    return [i for i in range(pagina.getCount()) if _eh_grafico(pagina.getByIndex(i))]


def _graficos_da_aba(aba):
    pagina = aba.DrawPage
    return [pagina.getByIndex(i) for i in _indices_de_graficos(aba)]


def _forma_atualizada(aba, forma):
    """Busca de novo a forma (mesmo gráfico) pelo nome interno. Em testes, uma referência que veio sem
    dados voltava a funcionar quando a forma era buscada outra vez."""
    try:
        persist = forma.PersistName
        pagina = aba.DrawPage
        for i in range(pagina.getCount()):
            f = pagina.getByIndex(i)
            if _eh_grafico(f) and f.PersistName == persist:
                return f
    except Exception:
        pass
    return forma


def _modelo_do_grafico(forma, aba=None, tentativas=6):
    """Documento do gráfico. O LibreOffice carrega objetos embutidos sob demanda e, de vez em quando,
    entrega a primeira referência sem dados. Tenta de novo buscando a forma outra vez e forçando o
    carregamento pela lista de gráficos da aba, antes de desistir."""
    for tentativa in range(tentativas):
        for atributo in ("Model", "EmbeddedObject"):
            try:
                m = getattr(forma, atributo)
                if m is not None:
                    return m
            except Exception:
                pass
        if aba is not None:
            try:
                m = aba.Charts.getByName(forma.PersistName).EmbeddedObject
                if m is not None:
                    return m
            except Exception:
                pass
            forma = _forma_atualizada(aba, forma)             # referência nova do mesmo gráfico
        time.sleep(0.3)
    raise ErroUsuario("Não consegui acessar os dados internos do gráfico.")


def _alvos(doc, todos=False):
    """Lista de (aba, obter_forma). A forma só é buscada na hora de usar: o LibreOffice descarrega
    gráficos embutidos que ficaram parados, e uma referência guardada antes pode vir sem dados."""
    if todos:
        resultado = []
        for i in range(doc.Sheets.Count):
            aba = doc.Sheets.getByIndex(i)
            for k in _indices_de_graficos(aba):
                resultado.append((aba, lambda a=aba, k=k: a.DrawPage.getByIndex(k)))
        if not resultado:
            raise ErroUsuario("Não há gráficos neste arquivo.")
        return resultado
    aba_ativa = doc.CurrentController.ActiveSheet
    selecionadas = _formas_selecionadas(doc)
    if selecionadas:
        return [(aba_ativa, lambda f=f: f) for f in selecionadas]
    indices = _indices_de_graficos(aba_ativa)
    if len(indices) == 1:
        return [(aba_ativa, lambda: aba_ativa.DrawPage.getByIndex(indices[0]))]
    if not indices:
        raise ErroUsuario("Não há gráfico nesta aba.")
    raise ErroUsuario("Há %d gráficos nesta aba. Clique uma vez no gráfico desejado e rode a macro de novo." % len(indices))


# ---------------------------------------------------------------- utilidades de formatação

def _tentar(avisos, rotulo, funcao):
    """Roda uma etapa isolada; falha vira aviso em vez de interromper o resto."""
    try:
        funcao()
    except Exception as e:
        avisos.append("%s: %s" % (rotulo, e))


def _definir(obj, **props):
    """Define propriedades uma a uma, ignorando as que o objeto não aceita."""
    for nome, valor in props.items():
        try:
            setattr(obj, nome, valor)
        except Exception:
            pass


def _estilo_texto(obj, tamanho=TAM_TEXTO, negrito=None, italico=None):
    _definir(obj, CharFontName=FONTE, CharFontNameAsian=FONTE, CharFontNameComplex=FONTE,
             CharHeight=float(tamanho), CharColor=PRETO)
    if negrito is not None:
        _definir(obj, CharWeight=150.0 if negrito else 100.0)
    if italico is not None:
        _definir(obj, CharPosture=uno.Enum("com.sun.star.awt.FontSlant", "ITALIC" if italico else "NONE"))


def _series(modelo):
    """Todas as séries do gráfico (API chart2)."""
    diagrama = modelo.getFirstDiagram()
    for sistema in diagrama.getCoordinateSystems():
        for tipo in sistema.getChartTypes():
            for serie in tipo.getDataSeries():
                yield serie


def _formas_internas(modelo):
    """Caixas de texto desenhadas dentro do gráfico (as 'chart user shapes' do OOXML)."""
    for metodo in ("getDrawPage", "getAdditionalShapes"):
        try:
            pagina = getattr(modelo, metodo)()
        except Exception:
            continue
        if pagina is not None:
            return pagina
    raise ErroUsuario("o gráfico não expõe as formas internas")


def _texto_forma(forma):
    try:
        return str(forma.String)
    except Exception:
        return ""


def _eh_caixa_fonte(forma):
    try:
        nome = str(forma.Name)
    except Exception:
        nome = ""
    return nome == NOME_CAIXA_FONTE or _texto_forma(forma).strip().lower().startswith(PREFIXOS_FONTE)


# ---------------------------------------------------------------- formatação (Graficos_Fonte)

def _linha(obj, cor, espessura=None):
    _definir(obj, LineStyle=uno.Enum("com.sun.star.drawing.LineStyle", "SOLID"), LineColor=cor)
    if espessura is not None:
        _definir(obj, LineWidth=espessura)


def _zero_na_escala(modelo):
    """True se o zero está estritamente dentro da escala do eixo Y (há valores negativos e positivos)."""
    ya = modelo.Diagram.getYAxis()
    return ya.Min < 0 < ya.Max


def _eh_linha_zero(eixo):
    """Reconhece o eixo X secundário criado por este script (cruza em 0 e não mostra rótulos)."""
    try:
        posicao = eixo.CrossoverPosition.value
        return (posicao == "ZERO" or (posicao == "VALUE" and eixo.CrossoverValue == 0)) and not eixo.DisplayLabels
    except Exception:
        return False


def _rotulo_no_ultimo_ponto(modelo):
    """True se alguma série mostra rótulo de dados (na série toda ou só no último ponto)."""
    def mostra(rotulo):
        return bool(rotulo.ShowNumber or rotulo.ShowCategoryName or rotulo.ShowCustomLabel)
    try:
        for serie in _series(modelo):
            if mostra(serie.Label):
                return True
            for par in serie.getDataSequences():
                valores = par.getValues()
                if valores is not None and "values-y" in str(valores.Role):
                    n = len(valores.getData())
                    if n and mostra(serie.getDataPointByIndex(n - 1).Label):
                        return True
    except Exception:
        pass
    return False


def _tem_subtitulo(modelo):
    try:
        return bool(modelo.HasSubTitle) and bool(str(modelo.SubTitle.String).strip())
    except Exception:
        return False


def formatar_grafico(forma, modelo, avisos):
    """Aplica o padrão RAF. 'avisos' recebe o que falhar."""
    estado = {"travado": False}

    def travar():
        if CONGELAR_ESTILOS and not estado["travado"]:
            try:
                modelo.lockControllers()
                estado["travado"] = True
            except Exception:
                pass

    def destravar():
        if estado["travado"]:
            try:
                modelo.unlockControllers()
            except Exception:
                pass
            estado["travado"] = False

    travar()
    try:
        _formatar_grafico(forma, modelo, avisos, destravar)
    finally:
        destravar()


def _formatar_grafico(forma, modelo, avisos, destravar):
    def tamanho():
        sz = forma.Size
        sz.Width = max(sz.Width, LARGURA_MIN)
        sz.Height = max(sz.Height, ALTURA_MIN)
        forma.Size = sz
    _tentar(avisos, "tamanho mínimo", tamanho)
    largura, altura = forma.Size.Width, forma.Size.Height

    def fundo():
        _definir(modelo.Area, FillStyle=uno.Enum("com.sun.star.drawing.FillStyle", "SOLID"),
                 FillColor=COR_FUNDO, FillTransparence=0)
    _tentar(avisos, "fundo do gráfico", fundo)

    def parede():
        w = modelo.Diagram.Wall
        if PAREDE_SEM_BORDA:
            _definir(w, LineStyle=uno.Enum("com.sun.star.drawing.LineStyle", "NONE"))
        if PAREDE_COR is not None:
            _definir(w, FillStyle=uno.Enum("com.sun.star.drawing.FillStyle", "SOLID"),
                     FillColor=PAREDE_COR, FillTransparence=0)
    _tentar(avisos, "parede do gráfico", parede)

    def textos():
        diagrama = modelo.Diagram
        for tem, titulo in (("HasXAxisTitle", "getXAxisTitle"), ("HasYAxisTitle", "getYAxisTitle"),
                            ("HasSecondaryXAxisTitle", "getSecondaryXAxisTitle"),
                            ("HasSecondaryYAxisTitle", "getSecondaryYAxisTitle")):
            try:
                if getattr(diagrama, tem):
                    _estilo_texto(getattr(diagrama, titulo)())
            except Exception:
                pass
        for serie in _series(modelo):           # rótulos de dados herdam da série
            _estilo_texto(serie)
    _tentar(avisos, "fontes (títulos de eixo, rótulos)", textos)

    def eixos():
        diagrama = modelo.Diagram
        for metodo in ("getXAxis", "getYAxis", "getSecondaryXAxis", "getSecondaryYAxis"):
            try:
                eixo = getattr(diagrama, metodo)()
            except Exception:
                continue
            if eixo is None:
                continue
            _estilo_texto(eixo)
            # marcas PRINCIPAIS: externas (2) e desenhadas no eixo, não junto dos rótulos; as secundárias não são alteradas
            _definir(eixo, Marks=2,
                     MarkPosition=uno.Enum("com.sun.star.chart.ChartAxisMarkPosition", "AT_AXIS"))
            _linha(eixo, PRETO)
    _tentar(avisos, "eixos", eixos)

    zero = [False]                                             # True se a linha em y = 0 se aplica a este gráfico

    def eixo_x():
        d = modelo.Diagram
        zero[0] = bool(LINHA_ZERO) and _zero_na_escala(modelo)
        x = d.getXAxis()
        if zero[0] and LINHA_ZERO_MODO == "EIXO":
            # o próprio eixo X vira a linha em y = 0 (preta), e os rótulos continuam embaixo
            x.CrossoverPosition = uno.Enum("com.sun.star.chart.ChartAxisPosition", "ZERO")
            x.LabelPosition = uno.Enum("com.sun.star.chart.ChartAxisLabelPosition", "OUTSIDE_START")
            _linha(x, COR_LINHA_ZERO, ESPESSURA_LINHA_ZERO)
        else:
            # eixo X cruzando o eixo Y no início (embaixo), mesmo com valores negativos
            x.CrossoverPosition = uno.Enum("com.sun.star.chart.ChartAxisPosition", "START")
            x.LabelPosition = uno.Enum("com.sun.star.chart.ChartAxisLabelPosition", "NEAR_AXIS")
    _tentar(avisos, "eixo X (posição e linha em y = 0)", eixo_x)

    def linha_zero_secundaria():
        # Modo "SECUNDARIO": eixo X secundário cruzando o Y no valor 0, sem rótulos nem marcas.
        d = modelo.Diagram
        existe = bool(d.HasSecondaryXAxis)
        if zero[0] and LINHA_ZERO_MODO == "SECUNDARIO":
            if existe and not _eh_linha_zero(d.getSecondaryXAxis()):
                raise ErroUsuario("o gráfico já usa o eixo X secundário; não criei a linha em y = 0")
            d.HasSecondaryXAxis = True
            ax = d.getSecondaryXAxis()
            ax.CrossoverPosition = uno.Enum("com.sun.star.chart.ChartAxisPosition", "VALUE")
            ax.CrossoverValue = 0.0
            ax.DisplayLabels = False
            ax.Marks = 0
            _linha(ax, COR_LINHA_ZERO, ESPESSURA_LINHA_ZERO)
        elif existe and _eh_linha_zero(d.getSecondaryXAxis()):
            d.HasSecondaryXAxis = False                         # modo trocado ou zero fora da escala: remove a linha criada aqui
    _tentar(avisos, "linha secundária em y = 0", linha_zero_secundaria)

    def grades():
        d = modelo.Diagram
        d.HasYAxisGrid = True                              # grade principal do eixo Y
        _linha(d.YMainGrid, COR_GRADE, ESPESSURA_GRADE)
        try:
            if d.HasXAxisGrid:                             # a do X só é formatada se já existir
                _linha(d.XMainGrid, COR_GRADE, ESPESSURA_GRADE)
        except Exception:
            pass
    _tentar(avisos, "grades", grades)

    destravar()    # a partir daqui o script LÊ medidas (título, legenda...): o gráfico precisa estar atualizado

    # ---- layout vertical dinâmico ------------------------------------------------------------
    #   topo:    título -> (subtítulo, se houver) -> área de plotagem
    #   rodapé:  área de plotagem -> (legenda, se houver, embaixo) -> caixa "Fonte"
    def estilos_do_cabecalho():
        if getattr(modelo, "HasMainTitle", False):
            t = modelo.Title
            _estilo_texto(t, TAM_TITULO, negrito=True)
            try:
                t.String = t.String.upper()                    # caixa-alta
            except Exception:
                pass
        if _tem_subtitulo(modelo):
            _estilo_texto(modelo.SubTitle, TAM_SUBTITULO, negrito=False)
            _definir(modelo.SubTitle, CharColor=COR_SUBTITULO)
        if getattr(modelo, "HasLegend", False):
            _estilo_texto(modelo.Legend)
            modelo.Legend.Alignment = uno.Enum("com.sun.star.chart.ChartLegendPosition", POSICAO_LEGENDA)
    _tentar(avisos, "fontes do título, subtítulo e legenda", estilos_do_cabecalho)

    def posicionar():
        y_topo = 0
        if getattr(modelo, "HasMainTitle", False):
            t = modelo.Title
            tam, pos = t.Size, t.Position
            pos.X, pos.Y = int(max(0, (largura - tam.Width) / 2)), 0
            t.Position = pos                                   # centralizado, no topo
            y_topo = tam.Height
        if _tem_subtitulo(modelo):
            st = modelo.SubTitle
            tam, pos = st.Size, st.Position
            pos.X, pos.Y = int(max(0, (largura - tam.Width) / 2)), int(y_topo + FOLGA_SUBTITULO)
            st.Position = pos                                  # logo abaixo do título
            y_topo += FOLGA_SUBTITULO + tam.Height
        y_base = altura - ALTURA_FONTE                         # a caixa "Fonte" ocupa o rodapé
        if getattr(modelo, "HasLegend", False) and POSICAO_LEGENDA == "BOTTOM":
            leg = modelo.Legend
            if LEGENDA_LARGURA_TOTAL:                          # caixa da legenda = largura do gráfico
                dim = leg.Size
                dim.Width = largura
                leg.Size = dim
            tam, pos = leg.Size, leg.Position
            pos.X = int(max(0, (largura - tam.Width) / 2))
            pos.Y = int(max(0, y_base - FOLGA - tam.Height))   # logo acima da "Fonte"
            leg.Position = pos
            y_base = pos.Y
        d = modelo.Diagram                                     # área de plotagem (com os rótulos dos eixos)
        ret = d.calculateDiagramPositionIncludingAxes()
        ret.Y = int(y_topo + FOLGA)
        ret.Height = int(max(1000, y_base - FOLGA - ret.Y))
        margem = MARGEM_DIREITA_ROTULOS if _rotulo_no_ultimo_ponto(modelo) else MARGEM_DIREITA
        ret.Width = int(max(1000, min(ret.Width, largura - ret.X - margem)))
        d.setDiagramPositionIncludingAxes(ret)
    _tentar(avisos, "posição do título, subtítulo, legenda e área de plotagem", posicionar)

    def caixas_internas():
        _garantir_caixa_fonte(modelo, largura, altura)         # cria "Fonte: " se não houver
        pagina = _formas_internas(modelo)
        for i in range(pagina.getCount()):
            f = pagina.getByIndex(i)
            fonte = _eh_caixa_fonte(f)
            if fonte or not ANOTACOES_TAMANHO_PROPRIO:
                _estilo_texto(f, TAM_TEXTO, italico=True if fonte else None)
            else:                                              # anotação: só Calibri e preto, mantém o tamanho
                _definir(f, CharFontName=FONTE, CharFontNameAsian=FONTE, CharFontNameComplex=FONTE, CharColor=PRETO)
            _definir(f, TextVerticalAdjust=uno.Enum("com.sun.star.drawing.TextVerticalAdjust", "CENTER"))
            if fonte:                                          # embaixo, na largura do gráfico
                tam = f.Size
                tam.Width, tam.Height = largura, ALTURA_FONTE
                f.Size = tam
                pos = f.Position
                pos.X, pos.Y = 0, max(0, altura - ALTURA_FONTE)
                f.Position = pos
    _tentar(avisos, "caixas de texto internas", caixas_internas)


# ---------------------------------------------------------------- título, séries e "Fonte" (grafico_title)

PADRAO_INTERVALO = re.compile(r"^\$?(?P<aba>.+?)\.\$?(?P<col>[A-Z]+)\$?(?P<lin>\d+)(?::\$?[A-Z]+\$?\d+)?$")


def _celula_e_texto(celula):
    """True se a célula guarda texto (digitado ou resultado de fórmula); False se vazia, número ou data."""
    tipo = str(celula.Type.value) if hasattr(celula.Type, "value") else str(celula.Type)
    if tipo == "TEXT":
        return True
    if tipo == "FORMULA":
        return celula.FormulaResultType2 == 2          # FormulaResult.STRING
    return False


def _texto_da_celula(aba, endereco):
    """Texto da célula. Em branco -> None (ignorar). Número/data -> ErroUsuario (nunca vira título)."""
    celula = aba.getCellRangeByName(endereco)
    if _celula_e_texto(celula):
        texto = celula.String.strip()
        return texto or None
    if not celula.String.strip():
        return None
    raise ErroUsuario("%s contém um número ou data ('%s'), não texto; ignorei essa célula" % (endereco, celula.String))


def _referencia(representacao):
    """"$'Fig 01'.$B$8:$B$476" -> ("Fig 01", "B", 8); None se não reconhecer."""
    m = PADRAO_INTERVALO.match(representacao.strip())
    if not m:
        return None
    return m.group("aba").strip("'"), m.group("col"), int(m.group("lin"))


def _trocar_linha_do_intervalo(representacao, nova_linha):
    """"$'Fig 01'.$B$8:$B$476" -> "$'Fig 01'.$B$<nova_linha>" (None se não reconhecer)."""
    m = PADRAO_INTERVALO.match(representacao.strip())
    if not m:
        return None
    aba = m.group("aba")
    return "$%s.$%s$%d" % (aba, m.group("col"), nova_linha)


def _nomes_das_series(modelo, aba, linha_cabecalho, avisos):
    """Aponta o nome de cada série para a célula do cabeçalho, na coluna dos valores.

    Só mexe quando a célula do cabeçalho é um texto não vazio; caso contrário mantém o nome atual
    e avisa (evita o "rótulo errado": uma data ou número no lugar do nome).
    """
    provedor = modelo.getDataProvider()
    for serie in _series(modelo):
        for par in serie.getDataSequences():
            valores = par.getValues()
            if valores is None or "values-y" not in str(valores.Role):
                continue
            ref = _referencia(valores.getSourceRangeRepresentation())
            atual = par.getLabel()
            if ref is None or atual is None:
                continue
            nome_aba, coluna, _ = ref
            if nome_aba != aba.Name:
                continue                                       # dados em outra aba: não mexe
            celula = aba.getCellRangeByName("%s%d" % (coluna, linha_cabecalho))
            if not _celula_e_texto(celula) or not celula.String.strip():
                avisos.append("nome da série (coluna %s): %s%d não é um texto; mantive o nome atual"
                              % (coluna, coluna, linha_cabecalho))
                continue
            novo_intervalo = _trocar_linha_do_intervalo(valores.getSourceRangeRepresentation(), linha_cabecalho)
            if novo_intervalo is None or novo_intervalo == atual.getSourceRangeRepresentation():
                continue
            novo = provedor.createDataSequenceByRangeRepresentation(novo_intervalo)
            try:
                novo.Role = atual.Role                         # mantém o papel do rótulo original
            except Exception:
                pass
            par.setLabel(novo)


def _achar_texto_fonte(aba):
    """Texto de 'Fonte'/'Source': primeiro na coluna A; senão, a primeira célula da aba."""
    cursor = aba.createCursor()
    cursor.gotoEndOfUsedArea(False)
    ult = cursor.RangeAddress
    dados = aba.getCellRangeByPosition(0, 0, ult.EndColumn, ult.EndRow).getDataArray()
    primeira = None
    for linha in dados:
        for j, v in enumerate(linha):
            if isinstance(v, str) and v.strip().lower().startswith(PREFIXOS_FONTE):
                if j == 0:
                    return v.strip()
                if primeira is None:
                    primeira = v.strip()
    return primeira


def _garantir_caixa_fonte(modelo, largura, altura, texto=None):
    """Devolve a caixa "Fonte" do gráfico, criando-a se não existir.

    texto=None: caixa existente com texto fica como está; caixa nova ou vazia recebe TEXTO_FONTE_PADRAO.
    texto="...": define o texto (na existente ou na nova).
    """
    pagina = _formas_internas(modelo)
    for i in range(pagina.getCount()):
        f = pagina.getByIndex(i)
        if _eh_caixa_fonte(f):
            if texto is not None:
                f.String = texto
            elif not _texto_forma(f).strip():                  # caixa existente, mas vazia
                f.String = TEXTO_FONTE_PADRAO
            return f
    f = modelo.createInstance("com.sun.star.drawing.TextShape")
    pagina.add(f)
    f.Name = NOME_CAIXA_FONTE
    tam = f.Size
    tam.Width, tam.Height = largura, ALTURA_FONTE
    f.Size = tam
    pos = f.Position
    pos.X, pos.Y = 0, max(0, altura - ALTURA_FONTE)
    f.Position = pos
    f.String = TEXTO_FONTE_PADRAO if texto is None else texto
    _definir(f, FillStyle=uno.Enum("com.sun.star.drawing.FillStyle", "NONE"),
             LineStyle=uno.Enum("com.sun.star.drawing.LineStyle", "NONE"),
             TextLeftDistance=150)                       # o "F" não encosta na borda
    return f


def titulo_e_fonte(aba, forma, modelo, cfg, avisos):
    """Aplica título, subtítulo, nomes das séries e fonte. Devolve linhas descrevendo o que foi lido.

    cfg: {"titulo": "A3", "subtitulo": "A4" ou None, "fonte": "A5" ou None (None = procura 'Fonte'/'Source'),
          "cabecalho": 7}. Célula em branco = ignorada (o que já está no gráfico fica como está).
    """
    lidos = []

    def titulo():
        texto = _texto_da_celula(aba, cfg["titulo"])
        if texto is None:
            raise ErroUsuario("%s está em branco; título não alterado" % cfg["titulo"])
        modelo.HasMainTitle = True
        modelo.Title.String = texto
        lidos.append("Título (%s): %s" % (cfg["titulo"], texto))
    _tentar(avisos, "título (%s)" % cfg["titulo"], titulo)

    if cfg.get("subtitulo"):
        def subtitulo():
            texto = _texto_da_celula(aba, cfg["subtitulo"])
            if texto is None:
                return                                         # em branco: ignora
            modelo.HasSubTitle = True
            modelo.SubTitle.String = texto
            lidos.append("Subtítulo (%s): %s" % (cfg["subtitulo"], texto))
        _tentar(avisos, "subtítulo (%s)" % cfg["subtitulo"], subtitulo)

    _tentar(avisos, "nomes das séries (linha %d)" % cfg["cabecalho"],
            lambda: _nomes_das_series(modelo, aba, cfg["cabecalho"], avisos))

    def fonte():
        if cfg.get("fonte"):
            texto = _texto_da_celula(aba, cfg["fonte"])
        else:
            texto = _achar_texto_fonte(aba)
        if texto is None:
            _garantir_caixa_fonte(modelo, forma.Size.Width, forma.Size.Height)   # em branco: ignora
            return
        if PREFIXAR_FONTE and not texto.lower().startswith(PREFIXOS_FONTE):
            texto = TEXTO_FONTE_PADRAO + texto
        _garantir_caixa_fonte(modelo, forma.Size.Width, forma.Size.Height, texto)
        lidos.append("Fonte (%s): %s" % (cfg.get("fonte") or "busca na aba", texto))
    _tentar(avisos, "texto da fonte", fonte)
    return lidos


# ---------------------------------------------------------------- rótulos por mês (AddSelectiveDataLabels)

FORMATOS_DATA = ("%Y-%m-%d", "%d/%m/%Y", "%m/%Y", "%Y-%m", "%Y/%m")


def mes_da_categoria(valor):
    """Mês (1-12) de uma categoria que seja data (número de série do Calc ou texto); senão None."""
    if isinstance(valor, bool) or valor is None:
        return None
    if isinstance(valor, (int, float)):
        # números pequenos são anos (2013, 2014...), não datas: o menor serial aceito é 1910
        if valor < SERIAL_MINIMO_DATA:
            return None
        try:
            return (EPOCA + timedelta(days=int(valor))).month
        except (OverflowError, ValueError):
            return None
    texto = str(valor).strip()
    for formato in FORMATOS_DATA:
        try:
            return datetime.strptime(texto[:10] if "%d" in formato else texto, formato).month
        except ValueError:
            pass
    return None


def _categorias(modelo):
    sistema = modelo.getFirstDiagram().getCoordinateSystems()[0]
    cats = sistema.getAxisByDimension(0, 0).getScaleData().Categories
    if cats is None:
        raise ErroUsuario("o gráfico não tem categorias (eixo X)")
    return list(cats.getValues().getData())


def rotulos_por_mes(modelo, mes):
    """Rótulo (valor + categoria) só nos pontos cuja categoria cai no mês; remove os demais."""
    meses = [mes_da_categoria(c) for c in _categorias(modelo)]
    if all(m is None for m in meses):
        raise ErroUsuario("as categorias do eixo X não são datas; não dá para escolher um mês")
    marcados = [i for i, m in enumerate(meses) if m == mes]
    for serie in _series(modelo):
        serie.resetAllDataPoints()                       # remove rótulos individuais antigos
        vazio = uno.createUnoStruct("com.sun.star.chart2.DataPointLabel")
        serie.setPropertyValue("Label", vazio)           # série sem rótulo (tudo False)
        for i in marcados:
            rotulo = uno.createUnoStruct("com.sun.star.chart2.DataPointLabel")
            rotulo.ShowNumber = True
            rotulo.ShowCategoryName = True
            rotulo.ShowLegendSymbol = False
            serie.getDataPointByIndex(i).Label = rotulo     # o ponto aceita atribuição, não setPropertyValue
    return len(marcados)


# ---------------------------------------------------------------- macros

def _executar(nome, acao, todos=False):
    ctx, doc = _contexto()
    try:
        alvos = _alvos(doc, todos)
    except ErroUsuario as e:
        _mensagem(doc, nome, str(e))
        return
    avisos, lidos, feitos = [], [], 0
    inicio = time.time()
    for aba, obter_forma in alvos:
        parciais = []
        try:
            forma = obter_forma()
            modelo = _modelo_do_grafico(forma, aba)
            info = acao(aba, forma, modelo, parciais)
            lidos.extend("[%s] %s" % (aba.Name, linha) for linha in (info or []))
            feitos += 1
        except ErroUsuario as e:
            parciais.append(str(e))
        except Exception as e:
            parciais.append("%s: %s" % (type(e).__name__, e))
        avisos.extend("[%s] %s" % (aba.Name, p) for p in parciais)
    texto = "%d gráfico(s) formatado(s) em %.1f s." % (feitos, time.time() - inicio)
    if lidos:
        texto += "\n\nUsado:\n- " + "\n- ".join(lidos[:12])
    if avisos:
        texto += "\n\nAvisos:\n- " + "\n- ".join(avisos[:12])
        if len(avisos) > 12:
            texto += "\n(+%d)" % (len(avisos) - 12)
    _mensagem(doc, nome, texto)


def _so_formatar(aba, forma, modelo, avisos):
    formatar_grafico(forma, modelo, avisos)


def _com_titulo(cfg):
    def acao(aba, forma, modelo, avisos):
        lidos = titulo_e_fonte(aba, forma, modelo, cfg, avisos)
        formatar_grafico(forma, modelo, avisos)
        return lidos
    return acao


CFG_PADRAO = {"titulo": CELULA_TITULO, "subtitulo": CELULA_SUBTITULO, "fonte": CELULA_FONTE,
              "cabecalho": LINHA_CABECALHO}
CFG_ENGLISH = {"titulo": "A4", "subtitulo": None, "fonte": None, "cabecalho": 6}   # layout antigo: 'Source' pesquisada


def GraficosFonte(*args):
    """Graficos_Fonte: aplica o padrão de formatação ao gráfico selecionado."""
    _executar("Formatar gráfico", _so_formatar)


def GraficoTitulo(*args):
    """grafico_title: título (A3), subtítulo (A4), fonte (A5), nomes das séries (linha 7) e formatação."""
    _executar("Título, fonte e formato", _com_titulo(CFG_PADRAO))


def GraficoTituloEnglish(*args):
    """grafico_title_english: título (A4), séries na linha 6, 'Source' pesquisada na aba."""
    _executar("Title, source and format", _com_titulo(CFG_ENGLISH))


def FormatarTodosGraficos(*args):
    """Aplica GraficosFonte a todos os gráficos do arquivo."""
    _executar("Formatar todos os gráficos", _so_formatar, todos=True)


def TituloETodosGraficos(*args):
    """Aplica GraficoTitulo a todos os gráficos do arquivo."""
    _executar("Título e formato em todos", _com_titulo(CFG_PADRAO), todos=True)


def RotulosPorMes(*args):
    """chart_format: rótulos de dados apenas nos pontos do mês escolhido (1-12)."""
    ctx, doc = _contexto()
    try:
        aba, obter_forma = _alvos(doc)[0]
        forma = obter_forma()
    except ErroUsuario as e:
        _mensagem(doc, "Rótulos por mês", str(e))
        return
    texto = _caixa_entrada("Rótulos por mês", "Digite o número do mês (1-12) para os rótulos de dados:", "12")
    if texto is None:
        return
    try:
        mes = int(texto)
        if not 1 <= mes <= 12:
            raise ValueError
    except ValueError:
        _mensagem(doc, "Rótulos por mês", "Mês inválido. Digite um número de 1 a 12.")
        return
    try:
        n = rotulos_por_mes(_modelo_do_grafico(forma, aba), mes)
        _mensagem(doc, "Rótulos por mês", "Rótulos adicionados em %d ponto(s) por série, para o mês %d." % (n, mes))
    except ErroUsuario as e:
        _mensagem(doc, "Rótulos por mês", str(e))
    except Exception as e:
        _mensagem(doc, "Rótulos por mês", "Falhou: %s: %s" % (type(e).__name__, e))


g_exportedScripts = (GraficosFonte, GraficoTitulo, GraficoTituloEnglish, RotulosPorMes,
                     FormatarTodosGraficos, TituloETodosGraficos)
