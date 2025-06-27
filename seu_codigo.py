import os
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl import load_workbook
from openpyxl.comments import Comment

def ler_jira(caminho_jira, colunas_esperadas):
    if not os.path.exists(caminho_jira):
        print("⚠️ Arquivo Jira.xlsx não encontrado, ignorando dados do Jira.")
        return None

    try:
        df_jira = pd.read_excel(caminho_jira, sheet_name='Your Jira Issues', header=0)
        if all(col in df_jira.columns for col in colunas_esperadas):
            return df_jira[colunas_esperadas]
        else:
            print("⚠️ Colunas esperadas não encontradas na aba 'Your Jira Issues', ignorando Jira.")
            return None
    except Exception as e:
        print(f"⚠️ Erro ao ler Jira.xlsx: {e}. Ignorando dados do Jira.")
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
        "Jira": {
            "header": "8989EB",
            "even": "C4C3F7",
            "odd":  "E8E7FC",
        },
        "Maximo": {
            "header": "8BC34A",
            "even": "FFFFFF",
            "odd":  "EEF7E3",
        },
        "Participantes": {
            "header": "4B7BEC",
            "even": "E6F0FF",
            "odd":  "FFFFFF",
        },
        "Verificação": {
            "header": "FF9900",
            "even": "FFF2CC",
            "odd": "FFE699"
        }
    }

    comentarios_participantes = {
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

        if aba in ["Jira", "Maximo"]:
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
                        cell.alignment = Alignment(horizontal='center' if j==1 else 'left', vertical='center', wrap_text=True, indent=1 if j != 1 else 0)
                        if i % 2 == 0:
                            cell.fill = fill_even
                        else:
                            cell.fill = fill_odd
                        if cell.column_letter in ['F', 'G']:
                            cell.number_format = "DD/MM/YYYY HH:mm"

            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                ws.column_dimensions[column].width = max(max_length + 2, 10)

            ws.sheet_view.showGridLines = True

        elif aba == "Participantes":
            # Limpa planilha Participantes
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

            # Escreve tabela
            for i, row_data in enumerate(participantes_obrigatorios):
                for j, value in enumerate(row_data):
                    cell = ws.cell(row=i + 1, column=j + 1, value=value)
                    if i == 0:
                        cell.fill = fill_header
                        cell.font = fonte_titulo
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    else:
                        cell.fill = fill_odd if i % 2 else fill_even
                        cell.font = fonte_normal
                        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True, indent=1 if j==0 else 0)
                    cell.border = thin_border

            # Comentários Participantes
            for coord, texto in comentarios_participantes.items():
                cell = ws[coord]
                cell.comment = Comment(texto, "GPT")

            # Ajusta largura colunas
            max_colunas = max(len(participantes_obrigatorios[0]), 2)
            for col_idx in range(1, max_colunas + 1):
                col_letter = get_column_letter(col_idx)
                max_length = 0
                for row_idx in range(1, len(participantes_obrigatorios) + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max(max_length + 5, 15)

            ws.sheet_view.showGridLines = True

        elif aba == "Verificação":
            # Modelo diferente para a aba Verificação
            # Limpa planilha
            for row in ws.iter_rows():
                for cell in row:
                    cell.value = None
                    cell.fill = PatternFill(fill_type=None)
                    cell.border = Border()
                    cell.font = Font()
                    cell.alignment = Alignment()

            # Cabeçalho customizado
            headers = ["Chave", "Resumo", "Status", "Descrição", "Relator", "Planned start date", "Planned end date"]
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_idx, value=header)
                cell.fill = fill_header
                cell.font = fonte_titulo
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = thin_border

            ws.sheet_view.showGridLines = True

def gerar_planilha_verificacao(df_jira, df_maximo, nome_saida, colunas_finais):
    sistemas = [
        "ARS\\NCR", "Athena", "Concentrador Fiscal", "Concsitef", "CTF", "Gescom", "Gold",
        "Guepardo", "MasterSaf", "Pegasus Descontos Comerciais", "SAD Contábil", "SAP",
        "SCE", "Sitef", "Storex", "TPLinux", "XRT"
    ]

    def filtro_sistema(texto):
        if pd.isna(texto):
            return False
        for sistema in sistemas:
            if sistema.lower() in str(texto).lower():
                return True
        return False

    import calendar
    from datetime import timedelta

    df_verificacao = pd.DataFrame(columns=colunas_finais)

    for df in [df_jira, df_maximo]:
        if df is None:
            continue

        for idx, row in df.iterrows():
            if filtro_sistema(row["Resumo"]) or filtro_sistema(row["Descrição"]):
                # Verifica datas Planned start date e Planned end date
                for col_data in ["Planned start date", "Planned end date"]:
                    data = row.get(col_data)
                    if pd.isna(data):
                        continue
                    ultimo_dia_mes = data.replace(day=calendar.monthrange(data.year, data.month)[1])
                    primeiro_dia_proximo_mes = (ultimo_dia_mes + timedelta(days=1))

                    if data.date() in [ultimo_dia_mes.date(), primeiro_dia_proximo_mes.date()] or \
                       (data.date() >= ultimo_dia_mes.date() and data.date() <= primeiro_dia_proximo_mes.date()):
                        df_verificacao = pd.concat([df_verificacao, pd.DataFrame([row[colunas_finais]])], ignore_index=True)
                        break

    if df_verificacao.empty:
        return None

    return df_verificacao

def padronizar_e_gerar_planilha(nome_saida='planilha_final.xlsx'):
    caminho_pasta = os.path.join(os.getcwd(), "arquivos")
    caminho_jira = os.path.join(caminho_pasta, "Jira.xlsx")

    colunas_finais = [
        "Chave", "Resumo", "Status", "Descrição", "Relator",
        "Planned start date", "Planned end date"
    ]

    df_jira = ler_jira(caminho_jira, colunas_finais)
    df_maximo = ler_maximo(caminho_pasta, colunas_finais)

    abas_criadas = []

    with pd.ExcelWriter(os.path.join(caminho_pasta, nome_saida), engine='openpyxl') as writer:
        if df_jira is not None:
            df_jira.to_excel(writer, sheet_name='Jira', index=False)
            abas_criadas.append('Jira')
        else:
            print("⚠️ Dados Jira não foram encontrados ou estão inválidos. Aba 'Jira' não será criada.")

        df_maximo.to_excel(writer, sheet_name='Maximo', index=False)
        abas_criadas.append('Maximo')

        # Aba Participantes com conteúdo e comentários
        import numpy as np
        df_dummy = pd.DataFrame(np.nan, index=range(10), columns=range(2))
        df_dummy.to_excel(writer, sheet_name='Participantes', index=False, header=False)
        abas_criadas.append('Participantes')

        # Gera planilha Verificação (apenas se tiver dados válidos)
        df_verificacao = gerar_planilha_verificacao(df_jira, df_maximo, nome_saida, colunas_finais)
        if df_verificacao is not None and not df_verificacao.empty:
            df_verificacao.to_excel(writer, sheet_name='Verificação', index=False)
            abas_criadas.append('Verificação')
        else:
            print("Nenhuma alteração conflitante com fechamento contábil, aba 'Verificação' não será criada.")

    aplicar_formatacao_excel(os.path.join(caminho_pasta, nome_saida), abas_criadas)

if __name__ == "__main__":
    padronizar_e_gerar_planilha()
