"""
Sou professor em um curso e gostaria de criar uma automação para emitir os certificados.
Os dados estão em uma planilha e gostaria de transferir para a imagem salvado com o formato jpg.
Os campos a serem preenchidos serão:

Nome do curso
Nome do participante
Tipo de participação
Data do início
Data do término
Data da emisão do certificado
Carga horária

1 PEGAR DADOS DA PLANILHA

2 TRANSFERIR PARA A IMAGEM DO CERTIFICADO


"""


import openpyxl
from PIL import Image, ImageDraw, ImageFont

# 1 PEGAR DADOS DA PLANILHA

# Abrir a Planilha
workbook_member = openpyxl.load_workbook('planilha_member.xlsx')
sheet_member = workbook_member['Sheet1']

for indice, line in enumerate(sheet_member.iter_rows(min_row=2)):
    
    nome_curso = line[0].value
    nome_particiante = line[1].value
    tipo_participacao = line[2].value
    data_inicio = line[3].value
    data_final = line[4].value
    carga_horaria = line[5].value
    data_emissao = line[6].value
  
    
# 2 TRANSFERIR PARA A IMAGEM DO CERTIFICADO 

# Definir a fonte
    font_bold = ImageFont.truetype('./tahomabd.ttf',90)
    font_default = ImageFont.truetype('./tahoma.ttf',80)
    font_date = ImageFont.truetype('./tahoma.ttf',55)

# Buscando e alterando a imagem
    image = Image.open('./cert_default.jpg')
    draw_image = ImageDraw.Draw(image)

# Insere nome do participante 
    draw_image.text((1020,825), nome_particiante,fill='black',font=font_bold)
# Insere nome do curso
    draw_image.text((1060,952), nome_curso,fill='black',font=font_default)
# Tipo de participante
    draw_image.text((1435,1070), tipo_participacao,fill='black',font=font_default)
# Carga horária
    draw_image.text((1480,1188),str(carga_horaria),fill='black',font=font_default)
# Data inicial
    draw_image.text((750,1770),data_inicio,fill='blue',font=font_date)
# Data final 
    draw_image.text((750,1930),data_final,fill='blue',font=font_date)
# Data emissão
    draw_image.text((2220,1930),data_emissao,fill='blue',font=font_date)

    image.save(f'./{indice} {nome_particiante} certificate.jpg')

