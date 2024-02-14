"""Automatizar o Envio dos dados da Planilha para preencher os campos mutáveis no certificado padrão.
Nome do curso,nome do participante, tipo de participação, data de inicio, data do final, carga horária,
data da emissão, assinaturas do GESTOR Geral, do coordenador e do aluno."""

import openpyxl
from PIL import Image,ImageDraw, ImageFont

#Abrir a planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = str(linha[5].value)
    data_emissao = linha[6].value

    #Tranferir para a imagem do certificado
    #Definindo a fonte a ser usada
    fonte_nome = ImageFont.truetype('./tahomabd.ttf',90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf',80)
    fonte_data = ImageFont.truetype('./tahoma.ttf',70)

    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    desenhar.text((1005,823),nome_participante,fill='black',font=fonte_nome)
    desenhar.text((1055,952),nome_curso,fill='black',font=fonte_geral)
    desenhar.text((1422,1065),tipo_participacao,fill='black',font=fonte_geral)
    desenhar.text((1475,1185),carga_horaria,fill='black',font=fonte_geral)
    desenhar.text((710,1765),data_inicio,fill='black',font=fonte_data)
    desenhar.text((710,1915),data_final,fill='black',font=fonte_data)
    desenhar.text((2175,1915),data_emissao,fill='black',font=fonte_data)

    image.save(f'./{indice} {nome_participante} certificado.pdf')
