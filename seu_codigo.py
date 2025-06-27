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
    import traceback
    caminho_xlsx = os.path.join(caminho_pasta, "Maximo.xlsx")
    caminho_csv = os.path.join(caminho_pasta, "Maximo.csv")
    df_maximo = None

    try:
        if os.path.exists(caminho_xlsx):
            print(f"Tentando ler {caminho_xlsx}")
            df = pd.read_excel(caminho_xlsx, sheet_name='Maximo')
            print("Colunas disponíveis no Excel:", df.columns.tolist())
        elif os.path.exists(caminho_csv):
            print(f"Tentando ler {caminho_csv}")
            df = pd.read_csv(caminho_csv, encoding="utf-8", sep=",")
            print("Colunas disponíveis no CSV:", df.columns.tolist())
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
        print("Colunas após renomear:", df.columns.tolist())

        df["Planned start date"] = pd.to_datetime(df["Planned start date"], errors="coerce", dayfirst=True)
        df["Planned end date"] = pd.to_datetime(df["Planned end date"], errors="coerce", dayfirst=True)
        df = df[df["Chave"] != "String"]
        df = df[df["Status"] == "AUTH"]

        if all(col in df.columns for col in colunas_esperadas):
            df_maximo = df[colunas_esperadas]
        else:
            print("⚠️ Algumas colunas esperadas estão faltando após filtragem.")

    except Exception as e:
        print("Erro no ler_maximo:", e)
        print(traceback.format_exc())

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
        "Verificação": {"header": "8989EB", "even": "C4C3F7", "odd": "E8E7FC"},  # mesma cor do Jira para manter padrão
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

        # Aba Participantes - conteúdo original e formatação
        if aba == "Participantes":
            # Limpar tudo antes de escrever
            for row in ws.iter_rows():
                for cell in row:
                    cell.value = None
                    cell.fill = PatternFill(fill_type=None)
                    cell.border = Border()
                    cell.font = Font()
                    cell.alignment = Alignment()

            participantes_obrigatorios = [
                ["Cargo", "Responsável"],
                ["Arquitetura:", ""],
                ["Field:", ""],
                ["Governança:", ""],
                ["Infraestrutura:", ""],
                ["N1:", ""],
                ["Operação:", ""],
                ["Segurança:", ""],
                ["Telecom:", ""],
            ]

            # Escrever a tabela participantes_obrigatorios
            for i, row_data in enumerate(participantes_obrigatorios):
                for j, value in enumerate(row_data):
                    cell = ws.cell(row=i+1, column=j+1, value=value)
                    if i == 0:
                        cell.fill = fill_header
                        cell.font = fonte_titulo
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    else:
                        cell.fill = fill_odd if i % 2 else fill_even
                        cell.font = fonte_normal
                        alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                        if j == 0:
                            alignment = Alignment(horizontal='left', vertical='center', wrap_text=True, indent=1)
                        cell.alignment = alignment
                    cell.border = thin_border

            # Adicionar comentários na coluna A
            for coord, texto in comentarios.items():
                if coord in ws:
                    ws[coord].comment = Comment(texto, 'GPT')

            # Ajustar largura colunas Participantes
            max_colunas = max(len(participantes_obrigatorios[0]), 2)
            for col_idx in range(1, max_colunas + 1):
                max_length = 0
                col_letter = get_column_letter(col_idx)
                for row_idx in range(1, len(participantes_obrigatorios) + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                adjusted_width = min(max(max_length + 5, 15), 50)
                ws.column_dimensions[col_letter].width = adjusted_width

            ws.sheet_view.showGridLines = True

        else:
            # Para outras abas (Jira, Maximo, Verificação), aplicar formatação padrão
            for i, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row,
                                                min_col=1, max_col=ws.max_column), 1):
                for j, cell in enumerate(row, 1):
                    cell.border = thin_border
                    if i == 1:
                        cell.fill = fill_header
                        cell.font = fonte_titulo
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    else:
                        cell.font = fonte_normal
                        if j == 1:
                            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                        else:
                            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True, indent=1)
                        if i % 2 == 0:
                            cell.fill = fill_even
                        else:
                            cell.fill = fill_odd
                        if cell.column_letter in ['F', 'G']:
                            cell.number_format = "DD/MM/YYYY HH:mm"

            # Ajustar largura colunas com coluna A maior
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                if column == 'A':
                    adjusted_width = max(max_length + 5, 20)
                else:
                    adjusted_width = min(max(max_length + 2, 10), 50)
                ws.column_dimensions[column].width = adjusted_width

            ws.sheet_view.showGridLines = True

    wb.save(caminho_arquivo)

def verificar_sistemas_em_fechamento(df_jira, df_maximo, caminho_saida):
    if df_jira is None:
        df_jira = pd.DataFrame(columns=["Resumo", "Descrição", "Planned start date", "Planned end date"])

    df_comb = pd.concat([df_jira, df_maximo], ignore_index=True)

    def tem_sistema(texto):
        texto = str(texto).lower()
        return any(s.lower() in texto for s in sistemas_fechamento)

    # Função para verificar se a data está no fechamento contábil
    def esta_no_fechamento(data):
        if pd.isna(data):
            return False
        dia = data.day
        # Considere dias finais do mês e dia 1 do próximo mês
        return dia in [28, 29, 30, 31, 1]

    df_filtrado = df_comb[df_comb.apply(
        lambda row: (
            tem_sistema(row.get("Resumo", "")) or tem_sistema(row.get("Descrição", ""))
        ) and (
            esta_no_fechamento(row.get("Planned start date")) or
            esta_no_fechamento(row.get("Planned end date"))
        ), axis=1)]

    if df_filtrado.empty:
        # Nenhuma change em conflito, não cria planilha Verificação
        return

    # Criar aba Verificação
    wb = load_workbook(caminho_saida)
    if "Verificação" in wb.sheetnames:
        ws = wb["Verificação"]
        wb.remove(ws)
    ws = wb.create_sheet("Verificação")

    # Colocar cabeçalho
    cabecalhos = list(df_filtrado.columns)
    ws.append(cabecalhos)

    # Preencher linhas
    for _, row in df_filtrado.iterrows():
        ws.append([row.get(col, '') for col in cabecalhos])

    wb.save(caminho_saida)

    # Aplicar formatação padrão (igual Jira e Maximo)
    aplicar_formatacao_excel(caminho_saida, ["Verificação"])
