import openpyxl
from PIL import Image, ImageDraw, ImageFont

# pegar dados da planilha
workbook_alunos = openpyxl.load_workbook("planilha_alunos.xlsx")
sheet_alunos = workbook_alunos['Sheet1'] #pra saber precisa olhar a planilha

for i, linha in enumerate(sheet_alunos.iter_rows(min_row=2)): # é a partir da row 2 pq a 1 possui o cabeçalho
    # para teste com um pode ser usado o max_row=2
    
    # definindo as fontes
    fonte_bold = ImageFont.truetype('./tahomabd.ttf', 90)
    fonte_normal = ImageFont.truetype('./tahoma.ttf', 80)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 55)

    # Carregar a imagem
    imagem = Image.open("./certificado_padrao.jpg")
    desenhar = ImageDraw.Draw(imagem) # para conseguir escrever sobre ela

   
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horária = linha[5].value
    data_emissao = linha[6].value

 # transferir dados da planilha para a imagem
    desenhar.text((1020,826), nome_participante, fill='black', font= fonte_bold)
    desenhar.text((1075, 950), nome_curso, fill='black', font=fonte_normal)
    desenhar.text((1445, 1070), tipo_participacao, fill='black', font=fonte_normal)
    desenhar.text((1495, 1182), f'{carga_horária} horas', fill='black', font=fonte_normal)

    # tamanho_data_inicio = ImageDraw.textlength(data_inicio, fonte_normal) - não está funcionando
    desenhar.text((750, 1770), data_inicio, fill='#4287f5', font=fonte_data)
    desenhar.text((750, 1930), data_final, fill='#4287f5', font=fonte_data)

    desenhar.text((2220, 1930), data_emissao, fill='#4287f5', font=fonte_data)



    imagem.save(f'./{i} - {nome_participante}.png')