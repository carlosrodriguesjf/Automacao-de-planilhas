

# lista = []
# for i in range(3):
#     teste = input('digite: ')
#     lista.append(teste)
# print(lista)


import openpyxl

workbook = openpyxl.Workbook()

selecao = []

# sheet_instrumentos = workbook.create_sheet('instrumentos')

# del workbook['Sheet']

# sheet_instrumentos = workbook['instrumentos']

# sheet_atual = sheet_instrumentos.append(['NOME','IDADE'])

palavra = input('Digite os dados a serem adicionados a uma nova linha, separados por vírgula: ')

palavra_separada = palavra.split(',')

print(palavra_separada)





# sheet_instrumentos.append(['Carlos ','41'])


# workbook.save('teste.xlsx')
