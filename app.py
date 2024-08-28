import openpyxl
from PIL import Image, ImageDraw, ImageFont

# Carrega a planilha do Excel
workbook_funcionarios = openpyxl.load_workbook('dados.xlsx')
planilha_dados = workbook_funcionarios['Plan1']

# Itera pelas linhas da planilha, começando da segunda linha
for linha in planilha_dados.iter_rows(min_row=2):
    nome = linha[0].value
    cargo = linha[1].value
    ramal = linha[2].value
    
    # Carrega as fontes com os tamanhos específicos
    font_nome = ImageFont.truetype('./Fonte.ttf', 75) #Modelo de fonte a ser usada, tamanho da fonte a ser ajustado de acordo com o tamanho do modelo de assinatura
    font_cargo = ImageFont.truetype('./Fonte.ttf', 75) #Modelo de fonte a ser usada, tamanho da fonte a ser ajustado de acordo com o tamanho do modelo de assinatura
    font_ramal_label = ImageFont.truetype('./Fonte.ttf', 64) #Modelo de fonte a ser usada, tamanho da fonte a ser ajustado de acordo com o tamanho do modelo de assinatura
    font_ramal = ImageFont.truetype('./Fonte.ttf', 64) #Modelo de fonte a ser usada, tamanho da fonte a ser ajustado de acordo com o tamanho do modelo de assinatura
    
    # Carrega a imagem base
    image = Image.open('./assinatura_email.png') #Modelo para a assinatura de email
    desenhar = ImageDraw.Draw(image)
    
    # Desenha o nome na posição (800, 350)
    desenhar.text((800, 320), nome, fill='white', font=font_nome) #Coordenadas a serem ajustadas de acordo com o tamanho da assinatura de email
    
    # Desenha o cargo na posição (450, 210)
    desenhar.text((800, 420), cargo, fill='white', font=font_cargo) #Coordenadas a serem ajustadas de acordo com o tamanho da assinatura de email
    
    # Desenha a palavra "Ramal" na posição (475, 350)
    desenhar.text((850, 690), "Ramal", fill='white', font=font_ramal_label) #Coordenadas a serem ajustadas de acordo com o tamanho da assinatura de email
    
    # Desenha o número do ramal na posição (475, 390)
    desenhar.text((1060, 690), str(ramal), fill='white', font=font_ramal) #Coordenadas a serem ajustadas de acordo com o tamanho da assinatura de email
    
    # Salva a imagem com o nome do participante
    image.save(f'./assinaturas/{nome}_cartao.png')
