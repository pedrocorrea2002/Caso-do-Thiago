import pyexcel as px
from collections import OrderedDict

teste = px.get_sheet(file_name="planilha_teste.xls")
dictio = {}

# Estou usando OrderedDict ao invés do dicionário padrão simplesmente porque de outra forma o save_as está mudando a ordem das colunas
formatted_sheet = OrderedDict({"N° da Etiqueta":[],"dia":[],"mes":[],"cont":[],"":[]})

# Agrupando a planilha com base na coluna "cont" e pegando para o mesmo valor da coluna "cont" o registro que possuir o maior valor na 5º coluna
for num,row in enumerate(teste) :
    if num != 0 :
        if str(row[3]) in dictio:
            if dictio[str(row[3])][4] < row[4] :
                dictio[str(row[3])] = row
        else:
            dictio[str(row[3])] = row

# Reformatando o dicionário
# Da forma como estava antes a nova planilha ia ser criada com 
for num,row in enumerate(dictio) :
    formatted_sheet["N° da Etiqueta"].append(dictio[row][0])
    formatted_sheet["dia"].append(dictio[row][1])
    formatted_sheet["mes"].append(dictio[row][2])
    formatted_sheet["cont"].append(dictio[row][3])
    formatted_sheet[""].append(dictio[row][4])

print(formatted_sheet.keys())


px.save_as(adict=formatted_sheet, dest_file_name="planilha_final.xlsx")
