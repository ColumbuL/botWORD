import pyautogui
import time

from docx import Document
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml



def texto_com_realce(run):
    highlight = run._element.xpath('.//w:highlight')
    return bool(highlight)


def extrair_codigos_sem_realce(caminho_arquivo):
    doc = Document(caminho_arquivo)
    codigos = []

    for paragrafo in doc.paragraphs:
        if '*' in paragrafo.text:
            partes = paragrafo.text.split('*')
            for parte in partes:
                codigo = parte.strip()
                if codigo: 
                    tem_realce = False
                    for run in paragrafo.runs:
                        if codigo in run.text and texto_com_realce(run):
                            tem_realce = True
                            break
                    if not tem_realce:
                        codigos.append(codigo)

    return codigos


def colar_conteudo_na_tela(conteudo, x, y):
    pyautogui.moveTo(x, y)
    pyautogui.click()
    pyautogui.write(conteudo, interval=0.05)

def clicar_para_enviar(x, y):
    pyautogui.moveTo(x, y)
    pyautogui.click()


caminho_arquivo = "C:/Users/PICHAU/Desktop/Lucas/RAD Python/test.docx"

x, y = 500, 500  
codigos_extraidos = extrair_codigos_sem_realce(caminho_arquivo)

for codigo in codigos_extraidos:
    time.sleep(1)
    colar_conteudo_na_tela(codigo, x, y)
    y += 50  
