import uno
import urllib.request
import io
import pandas as pd
from datetime import datetime, timedelta

def AtualizarSGS(*args):
    """Atualiza séries SGS na aba 'Tabela_SGS' com base nos IDs listados em 'Parametros' e datas opcionais."""

    # Conecta no documento ativo
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()

    # Função para criar ou obter aba
    def get_or_create_sheet(doc, sheet_name):
        sheets = doc.Sheets
        if not sheets.hasByName(sheet_name):
            sheets.insertNewByName(sheet_name, len(sheets))
        return sheets.getByName(sheet_name)

    # Cria ou obtém abas
    sheet_params = get_or_create_sheet(doc, "Parametros")
    sheet_dest = get_or_create_sheet(doc, "Tabela_SGS")

    # Locale pt-BR
    locale = uno.createUnoStruct('com.sun.star.lang.Locale')
    locale.Language = 'pt'
    locale.Country = 'BR'

    # Formato numérico com 2 casas decimais
    number_formats = doc.NumberFormats
    num_key = number_formats.queryKey("#.##0,00", locale, True)
    if num_key == -1:
        num_key = number_formats.addNew("#.##0,00", locale)

    # Formato de data DD/MM/YYYY
    date_key = number_formats.queryKey("DD/MM/YYYY", locale, True)
    if date_key == -1:
        date_key = number_formats.addNew("DD/MM/YYYY", locale)

    # Função para escrever log em F2
    def log_msg(msg):
        col_idx = 5  # F
        row_idx = 1  # 2ª linha
        cell = sheet_params.getCellByPosition(col_idx, row_idx)
        cell.String = msg
        cell.CellBackColor = 0xD3D3D3  # cinza claro
        cell.CharHeight = 12
        cell.HoriJustify = 1
        cell.VertJustify = 1
        sheet_params.Columns[col_idx].OptimalWidth = True
        sheet_params.Rows[row_idx].OptimalHeight = True

    # Converte datas do Calc para datetime
    def excel_date_to_datetime(excel_date):
        if isinstance(excel_date, (float, int)) and excel_date != 0:
            return datetime(1899, 12, 30) + timedelta(days=excel_date)
        return None

    # Ajusta cabeçalhos da aba Parametros
    headers = ["ID da Série", "Descrição (opcional)", "Data Inicial (DD/MM/AAAA)",
               "Data Final (DD/MM/AAAA)", "", "Log"]
    for col, header in enumerate(headers):
        cell = sheet_params.getCellByPosition(col, 0)
        cell.String = header
        cell.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")

    # Lê IDs e datas
    series_info = []
    for row in range(1, 100):
        id_cell = sheet_params.getCellByPosition(0, row)   # A
        start_cell = sheet_params.getCellByPosition(2, row)  # C
        end_cell = sheet_params.getCellByPosition(3, row)    # D
        series_id = id_cell.String.strip()
        if series_id != "":
            start_val = start_cell.Value
            end_val = end_cell.Value
            start_date = excel_date_to_datetime(start_val).strftime("%d/%m/%Y") if start_val else ""
            end_date = excel_date_to_datetime(end_val).strftime("%d/%m/%Y") if end_val else ""
            series_info.append((series_id, start_date, end_date))

    if not series_info:
        log_msg("Nenhum ID informado em Parametros! (coluna A a partir da linha 2)")
        return

    # Faz download de cada série
    frames = []
    for sid, start_date, end_date in series_info:
        try:
            url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{sid}/dados?formato=csv"
            if start_date:
                url += f"&dataInicial={start_date}"
            if end_date:
                url += f"&dataFinal={end_date}"
            with urllib.request.urlopen(url) as resp:
                data = resp.read().decode("utf-8")
            df = pd.read_csv(io.StringIO(data), sep=";", decimal=",")
            df.columns = ["data", sid]
            df["data"] = pd.to_datetime(df["data"], dayfirst=True)
            frames.append(df)
        except Exception as e:
            log_msg(f"Erro ao baixar série {sid}: {e}")

    if not frames:
        log_msg("Nenhuma série baixada com sucesso.")
        return

    # Merge por data
    result = frames[0]
    for df in frames[1:]:
        result = pd.merge(result, df, on="data", how="outer")
    result = result.sort_values("data")

    # Limpa aba destino
    sheet_dest.getCellRangeByName("A1:Z10000").clearContents(7)

    # Escreve cabeçalhos Tabela_SGS
    for col_idx, sid in enumerate(result.columns):
        if col_idx == 0:
            cell = sheet_dest.getCellByPosition(col_idx, 0)
            cell.String = "Data"
            cell.NumberFormat = date_key
            cell.CellBackColor = 0xADD8E6
            cell.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
        else:
            description = None
            for row in range(1, 100):
                id_cell = sheet_params.getCellByPosition(0, row)
                desc_cell = sheet_params.getCellByPosition(1, row)
                if id_cell.String.strip() == sid:
                    description = desc_cell.String.strip()
                    break
            header_text = description if description else sid
            cell = sheet_dest.getCellByPosition(col_idx, 0)
            cell.String = header_text
            cell.CellBackColor = 0xADD8E6
            cell.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")

    # Escreve valores
    for row_idx, row in enumerate(result.itertuples(index=False), start=1):
        for col_idx, value in enumerate(row):
            cell = sheet_dest.getCellByPosition(col_idx, row_idx)
            if col_idx == 0:
                # Coluna de data como número
                dt = row[0] if isinstance(row[0], datetime) else datetime.strptime(row[0], "%d/%m/%Y")
                cell.Value = (dt - datetime(1899, 12, 30)).days
                cell.NumberFormat = date_key
            else:
                try:
                    if pd.isna(value) or value == "":
                        cell.Value = 0
                    else:
                        cell.Value = float(value)
                    cell.NumberFormat = num_key
                except:
                    cell.Value = 0
                    cell.NumberFormat = num_key

    # Log de atualização
    now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    log_msg(f"Arquivo atualizado em: {now}")
