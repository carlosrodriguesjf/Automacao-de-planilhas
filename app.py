# PROJETO 2 - MESTRES DA AUTOMAÇÃO - AUTOMAÇÃO DE PLANILHAS


import openpyxl

# Criação do workbook e da Sheet 

workbook = openpyxl.Workbook()
criar_pagina = 's'

# Tela de boas vindas e criação das planilhas do workbook
print('Bem-vindo ao gerador de planilhas!\nPara começar vamos criar uma nova página dentro de uma planilha')

while criar_pagina == 's':
    nome_pagina = input('Digite o nome da página: ')
    workbook.create_sheet(nome_pagina)
    criar_pagina = input('Criar mais uma página nesta planilha?(s/n): ')
print(workbook.sheetnames)    


pagina_manipulada = input('Digite o nome da página a ser manipulada: ')
sheet_atual = workbook[pagina_manipulada]









workbook.save('dados.xlsx')