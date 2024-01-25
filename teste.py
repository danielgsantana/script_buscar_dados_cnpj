from openpyxl import Workbook, load_workbook

planilha = Workbook()

aba_ativa = planilha.active

aba_ativa['A1'] = 2

planilha.save('info.xlsx')