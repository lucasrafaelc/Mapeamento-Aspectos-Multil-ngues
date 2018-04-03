# -*- coding: utf-8 -*-

from googletrans import Translator
import xml.etree.ElementTree as ET
import time
import xlwt

#languages = ["nl", "ru", "en", "es", "tr"]
languages = ["tr"]

clusters = {"RESTAURANT#GENERAL", 'RESTAURANT#PRICES', 'RESTAURANT#MISCELLANEOUS', 'FOOD#PRICES', 'FOOD#QUALITY', 'FOOD#STYLE_OPTIONS', 'DRINKS#PRICES', 'DRINKS#QUALITY', 'DRINKS#STYLE_OPTIONS', 'SERVICE#GENERAL', 'AMBIENCE#GENERAL', 'LOCATION#GENERAL'}

translator = Translator()

lista_aspectos = []

arquivo = "aspectos_"+languages[0]+".xls"
print(arquivo)
#Leitura dos Arquivos
for lang in languages:
    tree = ET.parse(lang+".xml")
    root = tree.getroot()
    for sentence in root.iter('sentence'):
        #texto = sentence.find('text').text.lower()
        for opinions in sentence.findall('Opinions'):
            for opinion in opinions.findall('Opinion'):
                target = opinion.attrib['target']
                #print(type(target))
                #print(target)
                #print(target, type(target))
                if target != 'NULL' and target != '':
                    target = target.lower()
                    #print(target, type(target))
                    if lang != 'en':
                        traducao = translator.translate(target, src=lang, dest='en').text
                        time.sleep( 0.25 )
                    else:
                        traducao = target
                    categoria = opinion.attrib['category']
                    #print(type(target))
                    repeticao = False
                    for aspectos in lista_aspectos:
                        if aspectos[0] == target and aspectos[2] == categoria:
                            aspectos[4] = aspectos[4] + 1
                            print(target.decode('cp1251'), " REPETIDO ", aspectos[4], " vezes")
                            repeticao = True
                    if not repeticao:
                        try:
                            print(type(target), type(traducao))
                            target = str(target.encode('utf8'))
                            traducao = str(traducao.encode('utf8'))
                            lista = [target, traducao, categoria, lang, 1]
                            print(str(target.encode('utf8')), lista)
                            print(type(target), type(traducao))
                        except UnicodeEncodeError:
                            print("/t/t/tERRO: ", target ,traducao, categoria, lang, 1)
                        lista_aspectos.append(lista)
print(lista_aspectos)

#Gravação dos dados em planilhas
wb = xlwt.Workbook(encoding="UTF-8")

for categoria in clusters:
    ws = wb.add_sheet(categoria)
    # Títulos das colunas
    titles = ["Aspecto","Tradução", "Idioma", "TF"]
    # Escrevendo títulos na primeira linha do arquivo
    for i in range(len(titles)):
        ws.write(0, i, titles[i])

    i = 1
    for aspecto in lista_aspectos:
        if aspecto[2] == categoria:
            # Escrevendo o identificar na 1ª coluna da linha i
            ws.write(i, 0, aspecto[0])
            ws.write(i, 1, aspecto[1])
            ws.write(i, 2, aspecto[3])
            ws.write(i, 3, aspecto[4])
            ws.write(i, 4, aspecto[2])
            i += 1
            
# Salvando
wb.save(arquivo)

