
import pandas as pd
import os
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl import load_workbook
from openpyxl.comments import Comment

sistemas_fechamento = [
    "ARS", "NCR", "Athena", "Concentrador Fiscal", "Concsitef", "CTF",
    "Gescom", "Gold", "Guepardo", "MasterSaf", "Pegasus Descontos Comerciais",
    "SAD Contábil", "SAP", "SCE", "Sitef", "Storex", "TPLinux", "XRT"
]

def ler_jira(caminho_jira, colunas_esperadas):
    if not os.path.exists(caminho_jira):
        return None
    try:
        df_jira = pd.read_excel(caminho_jira, sheet_name='Your Jira Issues', header=0)
        if all(col in df_jira.columns for col in colunas_esperadas):
            return df_jira[colunas_esperadas]
        else:
            return None
    except:
        return None

def ler_maximo(caminho_pasta, colunas_esperadas):
    caminho_xlsx = os.path.join(caminho_pasta, "Maximo.xlsx")
    caminho_csv = os.path.join(caminho_pasta, "Maximo.csv")
    df_maximo = None
    try:
        if os.path.exists(caminho_xlsx):
            df = pd.read_excel(caminho_xlsx, sheet_name='Maximo')
        elif os.path.exists(caminho_csv):
            df = pd.read_csv(caminho_csv, encoding="utf-8", sep=",")
        else:
            raise FileNotFoundError("Nenhum arquivo Maximo encontrado")

        df = df.rename(columns={
            "change_number": "Chave",
            "summary": "Resumo",
            "status": "Status",
            "details": "Descrição",
            "owner_name": "Relator",
            "schedule_start": "Planned start date",
            "schedule_finish": "Planned end date"
        })
        df["Planned start date"] = pd.to_datetime(df["Planned start date"], errors="coerce", dayfirst=True)
        df["Planned end date"] = pd.to_datetime(df["Planned end date"], errors="coerce", dayfirst=True)
        df = df[df["Chave"] != "String"]
        df = df[df["Status"] == "AUTH"]

        if all(col in df.columns for col in colunas_esperadas):
            df_maximo = df[colunas_esperadas]

    except:
        pass

    if df_maximo is None:
        raise FileNotFoundError("Erro ao processar o Maximo")
    return df_maximo

def aplicar_formatacao_excel(caminho_arquivo, abas):
    wb = load_workbook(caminho_arquivo)
    fonte_normal = Font(name="Montserrat", size=11)
    fonte_titulo = Font(name="Montserrat", size=12, bold=True)
    thin_side = Side(border_style="thin", color="C4C7C5")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    cores = {
        "Jira": {"header": "8989EB", "even": "C4C3F7", "odd": "E8E7FC"},
        "Maximo": {"header": "8BC34A", "even": "FFFFFF", "odd": "EEF7E3"},
        "Participantes": {"header": "4B7BEC", "even": "E6F0FF", "odd": "FFFFFF"},
        "Verificação": {"header": "FF8A65", "even": "FFF3EE", "odd": "FFE0D3"},
    }

    comentarios = {
        "A2": "Igor Campos , Reginaldo Tadashi , Newton Albuquerque , Adilson Bassani",
        "A3": "Clodoaldo Dias , Wesley Magalhães",
        "A4": "Roberto , Wagner",
        "A5": "Sergio Massao , Wilton Carvalho",
        "A6": "Ricardo Witsmiszyn",
        "A7": "Enrique Dias , Ricardo Witsmiszyn",
        "A8": "Thais Yuta , Mauricio Souza , Guilherme Perdroso",
        "A9": "Thiago Pezzini , Samuel Silva"
    }

    for aba in abas:
        if aba not in wb.sheetnames:
            continue
        ws = wb[aba]
        cfg = cores.get(aba, {})
        fill_header = PatternFill(start_color=cfg.get("header", "CCCCCC"), fill_type="solid")
        fill_even = PatternFill(start_color=cfg.get("even", "FFFFFF"), fill_type="solid")
        fill_odd = PatternFill(start_color=cfg.get("odd", "FFFFFF"), fill_type="solid")

        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column), 1):
            for j, cell in enumerate(row, 1):
                cell.border = thin_border
                if i == 1:
                    cell.fill = fill_header
                    cell.font = fonte_titulo
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                else:
                    cell.font = fonte_normal
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                    cell.fill = fill_even if i % 2 == 0 else fill_odd
                if cell.column_letter in ['F', 'G']:
                    cell.number_format = "DD/MM/YYYY HH:mm"

        if aba == "Participantes":
            for coord, text in comentarios.items():
                if coord in ws:
                    ws[coord].comment = Comment(text, 'GPT')

        for col in ws.columns:
            max_len = max((len(str(cell.value)) for cell in col if cell.value), default=0)
            width = min(max_len + 5, 50)
            ws.column_dimensions[col[0].column_letter].width = width

    wb.save(caminho_arquivo)

def verificar_sistemas_em_fechamento(df_jira, df_maximo, caminho_saida):
    if df_jira is None:
        df_jira = pd.DataFrame(columns=["Resumo", "Descrição", "Planned start date", "Planned end date"])

    df_comb = pd.concat([df_jira, df_maximo], ignore_index=True)

    def tem_sistema(texto):
        texto = str(texto).lower()
        return any(s.lower() in texto for s in sistemas_fechamento)

    df_filtrado = df_comb[df_comb.apply(
        lambda row: (
            tem_sistema(row.get("Resumo", "")) or tem_sistema(row.get("Descrição", ""))) and (
            (row.get("Planned start date", pd.NaT).day in [28, 29, 30, 31, 1] if pd.notna(row.get("Planned start date")) else False)
            or
            (row.get("Planned end date", pd.NaT).day in [28, 29, 30, 31, 1] if pd.notna(row.get("Planned end date")) else False)
        ), axis=1)]

    if not df_filtrado.empty:
        from openpyxl import load_workbook
        wb = load_workbook(caminho_saida)
        ws = wb.create_sheet("Verificação")

        cabecalhos = list(df_filtrado.columns)
        ws.append(cabecalhos)
        for i, row in df_filtrado.iterrows():
            ws.append([row.get(col, '') for col in cabecalhos])
        aplicar_formatacao_excel(caminho_saida, ["Verificação"])
        wb.save(caminho_saida)
