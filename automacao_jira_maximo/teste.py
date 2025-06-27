from openpyxl import Workbook
from openpyxl.comments import Comment

# Cria uma nova planilha
wb = Workbook()
ws = wb.active

# Adiciona um comentário na célula A1
cell = ws['A1']
comment = Comment("Este é um teste de comentário.", "GPT")
cell.comment = comment

# Salva a planilha
wb.save("teste_comentario.xlsx")

print("Arquivo teste_comentario.xlsx gerado com um comentário na célula A1.")