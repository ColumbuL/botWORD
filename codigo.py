import pyautogui
import time

from docx import Document
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml


# Função para verificar se o texto tem marcação de cor
def texto_com_realce(run):
    highlight = run._element.xpath('.//w:highlight')
    return bool(highlight)


# Função para extrair códigos específicos ignorando os com realce
def extrair_codigos_sem_realce(caminho_arquivo):
    doc = Document(caminho_arquivo)
    codigos = []

    for paragrafo in doc.paragraphs:
        # Verifica se o parágrafo contém asterisco
        if '*' in paragrafo.text:
            # Separa os textos pelos asteriscos
            partes = paragrafo.text.split('*')
            for parte in partes:
                codigo = parte.strip()
                if codigo:  # Adiciona se não for uma string vazia
                    # Verifica se algum dos runs tem realce
                    tem_realce = False
                    for run in paragrafo.runs:
                        if codigo in run.text and texto_com_realce(run):
                            tem_realce = True
                            break
                    if not tem_realce:
                        codigos.append(codigo)

    return codigos

# Função para colar o conteúdo em uma posição específica da tela
def colar_conteudo_na_tela(conteudo, x, y):
    # Mover o cursor para a posição x, y
    pyautogui.moveTo(x, y)

    # Clicar na posição para focar
    pyautogui.click()

    # Colar o conteúdo (Simula a combinação Ctrl+V)
    pyautogui.write(conteudo, interval=0.05)

# Função para clicar em uma posição específica para enviar o código
def clicar_para_enviar(x, y):
    pyautogui.moveTo(x, y)
    pyautogui.click()


# Caminho para o arquivo Word
caminho_arquivo = "C:/Users/PICHAU/Desktop/Lucas/RAD Python/test.docx"

# Coordenadas x e y onde você quer colar o conteúdo
x, y = 500, 500  # Ajuste conforme necessário

# Extrair os códigos
codigos_extraidos = extrair_codigos_sem_realce(caminho_arquivo)

# Loop para colar cada código em uma posição específica
for codigo in codigos_extraidos:
    # Espera um tempo antes de colar, se necessário
    time.sleep(1)

    # Colar o conteúdo na posição especificada
    colar_conteudo_na_tela(codigo, x, y)

    # Mover para a próxima posição ou fazer outra ação, se necessário
    # Por exemplo, movendo o cursor um pouco para baixo para o próximo código
    y += 50  # Ajuste conforme necessário para mover a posição de colagem