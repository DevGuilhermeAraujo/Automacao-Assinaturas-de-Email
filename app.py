import openpyxl
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

BASE_DIR = Path(__file__).resolve().parent
FONT_PATH = BASE_DIR / 'Fonte.otf'
TEMPLATE_PATH = BASE_DIR / 'assinatura_email.png'
EXCEL_PATH = BASE_DIR / 'dados.xlsx'
OUTPUT_DIR = BASE_DIR / 'assinaturas'

if not FONT_PATH.exists():
    raise FileNotFoundError(f'Arquivo de fonte não encontrado: {FONT_PATH}')

if not TEMPLATE_PATH.exists():
    raise FileNotFoundError(f'Arquivo de imagem não encontrado: {TEMPLATE_PATH}')

if not EXCEL_PATH.exists():
    raise FileNotFoundError(f'Arquivo do Excel não encontrado: {EXCEL_PATH}')

OUTPUT_DIR.mkdir(exist_ok=True)

# Carrega a planilha do Excel
workbook_funcionarios = openpyxl.load_workbook(EXCEL_PATH)
planilha_dados = workbook_funcionarios['Plan1']

# Carrega as fontes com os tamanhos específicos
font_nome = ImageFont.truetype(FONT_PATH, 75)
font_cargo = ImageFont.truetype(FONT_PATH, 75)
font_ramal_label = ImageFont.truetype(FONT_PATH, 64)
font_ramal = ImageFont.truetype(FONT_PATH, 64)

# Itera pelas linhas da planilha, começando da segunda linha
for linha in planilha_dados.iter_rows(min_row=2):
    nome = linha[0].value
    cargo = linha[1].value
    ramal = linha[2].value

    if not nome:
        continue

    # Carrega a imagem base
    image = Image.open(TEMPLATE_PATH)
    desenhar = ImageDraw.Draw(image)

    # Desenha o nome na posição (800, 320)
    desenhar.text((800, 320), str(nome), fill='white', font=font_nome)

    # Desenha o cargo na posição (800, 420)
    desenhar.text((800, 420), str(cargo or ''), fill='white', font=font_cargo)

    # Desenha a palavra "Ramal" na posição (850, 690)
    desenhar.text((850, 690), 'Ramal', fill='white', font=font_ramal_label)

    # Desenha o número do ramal na posição (1060, 690)
    desenhar.text((1060, 690), str(ramal or ''), fill='white', font=font_ramal)

    # Salva a imagem com o nome do participante
    output_path = OUTPUT_DIR / f'{nome}_cartao.png'
    image.save(output_path)

