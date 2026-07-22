"""Script auxiliar (não faz parte do programa) para gerar as imagens de
exemplo usadas no README. Rodar uma vez e descartar se quiser."""
import os
from PIL import Image, ImageDraw, ImageFont

BASE = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = r"C:\Windows\Fonts"


def fonte(nome, tamanho):
    return ImageFont.truetype(os.path.join(FONT_DIR, nome), tamanho)


F_REG = fonte("segoeui.ttf", 16)
F_BOLD = fonte("segoeuib.ttf", 16)
F_TITLE = fonte("segoeuib.ttf", 18)
F_SMALL = fonte("segoeui.ttf", 14)
F_MONO = fonte("consola.ttf", 18)
F_MONO_B = fonte("consolab.ttf", 18)

VERDE = "#1a7a3c"
VERDE_CLARO = "#e6f4ea"
CINZA_CLARO = "#f3f3f3"
CINZA_TXT = "#8a8a8a"
AZUL_HEADER = "#217346"  # verde-excel
BORDA = "#c9c9c9"


# ---------------------------------------------------------------------------
# Imagem 1: planilha de exemplo
# ---------------------------------------------------------------------------
colunas = [
    ("Atendimento", "646963", True),
    ("Data", "21/01/2026", False),
    ("Logradouro", "Rua Manuel Pereira Ninho", True),
    ("Nº", "80", True),
    ("Bairro", "Alcântara", True),
    ("Materiais Aplicados", "Cabo PP 2X1,5mm...", False),
]
linhas_dados = [
    ["646963", "21/01/2026", "Rua Manuel Pereira Ninho", "80", "Alcântara", "Cabo PP 2X1,5mm, Qtd:4.00"],
    ["646962", "21/01/2026", "Rua Walter da Costa Dias", "261", "Antonina", "Reparo"],
    ["646960", "21/01/2026", "Rua Maria de Souza", "12", "Cruzeiro do Sul", "Luminária LED 60W, Qtd:1.00"],
]

larguras = [110, 100, 260, 50, 150, 260]
largura_total = sum(larguras)
alt_titulo = 40
alt_header = 40
alt_linha = 36
margem = 20
legenda_h = 70

img_w = largura_total + margem * 2
img_h = alt_titulo + alt_header + alt_linha * len(linhas_dados) + margem * 2 + legenda_h

img = Image.new("RGB", (img_w, img_h), "white")
d = ImageDraw.Draw(img)

x0, y0 = margem, margem

# título (linha 1 da planilha real)
d.rectangle([x0, y0, x0 + largura_total, y0 + alt_titulo], fill="#f8f8f8", outline=BORDA)
d.text((x0 + 10, y0 + 10), "ATENDIMENTOS DE MANUTENÇÃO - SÃO GONÇALO  (linha de título — ignorada)", font=F_SMALL, fill=CINZA_TXT)
y = y0 + alt_titulo

# header
x = x0
for (nome, _exemplo, usada), larg in zip(colunas, larguras):
    cor_fundo = AZUL_HEADER if usada else "#9e9e9e"
    d.rectangle([x, y, x + larg, y + alt_header], fill=cor_fundo, outline=BORDA)
    d.text((x + 8, y + 10), nome, font=F_BOLD, fill="white")
    x += larg
y += alt_header

# linhas de dados
for i, linha in enumerate(linhas_dados):
    x = x0
    for valor, (nome, _ex, usada), larg in zip(linha, colunas, larguras):
        cor_fundo = VERDE_CLARO if usada else CINZA_CLARO
        d.rectangle([x, y, x + larg, y + alt_linha], fill=cor_fundo, outline=BORDA)
        texto = valor if len(valor) < 30 else valor[:27] + "..."
        d.text((x + 8, y + 8), texto, font=F_REG, fill="#222222")
        x += larg
    y += alt_linha

# legenda
y += 14
d.rectangle([x0, y, x0 + 20, y + 20], fill=VERDE_CLARO, outline=VERDE)
d.text((x0 + 28, y), "coluna que o programa usa (nomes flexíveis, ver README)", font=F_SMALL, fill="#222222")
y += 28
d.rectangle([x0, y, x0 + 20, y + 20], fill=CINZA_CLARO, outline=BORDA)
d.text((x0 + 28, y), "coluna ignorada pelo programa (pode existir ou não)", font=F_SMALL, fill="#222222")

img.save(os.path.join(BASE, "exemplo_planilha.png"))


# ---------------------------------------------------------------------------
# Imagem 2: estrutura de pastas de fotos
# ---------------------------------------------------------------------------
linhas = [
    ("fotos\\", False, None),
    ("├── 646963\\", True, "646963"),
    ("│    ├── foto1.jpg", False, None),
    ("│    └── foto2.jpg", False, None),
    ("├── 646962\\", True, "646962"),
    ("│    └── foto1.jpg", False, None),
    ("└── 646960\\", True, "646960"),
    ("     └── foto1.jpg", False, None),
]

pad = 20
line_h = 30
w2 = 560
h2 = pad * 2 + line_h * len(linhas) + 50

img2 = Image.new("RGB", (w2, h2), "white")
d2 = ImageDraw.Draw(img2)

y = pad
d2.text((pad, y), "pasta do pack de fotos (nome da pasta = número do Atendimento)", font=F_SMALL, fill=CINZA_TXT)
y += 34
for texto, destaque, numero in linhas:
    if destaque:
        # calcula onde o número começa para destacar só ele
        prefixo = texto.split(numero)[0]
        d2.text((pad, y), prefixo, font=F_MONO, fill="#222222")
        largura_prefixo = d2.textlength(prefixo, font=F_MONO)
        d2.text((pad + largura_prefixo, y), numero, font=F_MONO_B, fill=VERDE)
        resto = texto.split(numero)[1]
        largura_num = d2.textlength(numero, font=F_MONO_B)
        d2.text((pad + largura_prefixo + largura_num, y), resto, font=F_MONO, fill="#222222")
    else:
        d2.text((pad, y), texto, font=F_MONO, fill="#555555")
    y += line_h

img2.save(os.path.join(BASE, "exemplo_pastas.png"))

print("Imagens geradas em", BASE)
