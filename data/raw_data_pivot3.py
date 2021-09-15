# -*- coding: utf-8 -*-

#-----------------------------------------------------------------------------------------------------------
#-- ASSUNTO       : Extração dados pivot table diesel by fuel and type                                    --
#-- PERIODICIDADE : Sob demanda                                                                           --
#-- AUTOR         : Thiago Damião de Freitas                                                              --
#--                                                                                                       --
#-- DESCRICAO     : Script para extração dos dados referentes a pivot table 3 do arquivo vendas-combustí- --
#--                 -veis-m3.xls                                                                          --
#--                                                                                                       --
#-- TABELA FINAL  : diesel_by_fuel                                                                        --
#-- DT ALTERACAO  : 15/09/2021                                                                            --
#-----------------------------------------------------------------------------------------------------------

import win32com.client as win32
import os, sys
import zipfile
import xml.etree.ElementTree as et
import pandas as pd

def conv_to_xlsx(): #convert xls format to xlsx function
    file_name = 'exame_raizen/data/vendas-combustiveis-m3.xls'
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(file_name)
    wb.SaveAs(file_name+"x", FileFormat = 51)
    wb.Close()                               
    excel.Application.Quit()

def rename_xlsx_to_zip():
    path = 'exame_raizen/data/'
    dir = os.listdir(path)
    for file in dir:
        if file == 'vendas-combustiveis-m3.xlsx':
            os.rename('exame_raizen/data/vendas-combustiveis-m3.xlsx', 'exame_raizen/data/vendas-combustiveis-m3.zip')
            print('A extenção do arquivo', file, 'foi substituida para (.csv)!!!')

def unzip_file():
    path = 'exame_raizen/data/'
    dir = os.listdir(path)
    for file in dir:
        with zipfile.ZipFile('exame_raizen/data/vendas-combustiveis-m3.zip', 'r') as zip_ref:
            zip_ref.extractall('exame_raizen/unzip/')

def delete_zip_file():
    path = 'exame_raizen/data/'
    dir = os.listdir(path)
    for file in dir:
        if file == 'vendas-combustiveis-m3.zip':
            os.remove('exame_raizen/data/vendas-combustiveis-m3.zip')
            print('O arquivo', file, 'foi removido do diretório:', path)            

conv_to_xlsx()
rename_xlsx_to_zip()
unzip_file()
delete_zip_file()

definitions = 'exame_raizen/unzip/xl/pivotCache/pivotCacheDefinition3.xml'
defdict = {}
columnas = []
e = et.parse(definitions).getroot()
for fields in e.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}cacheFields'):
  for cidx, field in enumerate(fields.getchildren()):
      columna = field.attrib.get('name')
      defdict[cidx] = []
      columnas.append(columna)
      for value in field.getchildren()[0].getchildren():
          tagname = value.tag
          defdict[cidx].append(value.attrib.get('v', 0))

dfdata = []
bdata = 'exame_raizen/unzip/xl/pivotCache/pivotCacheRecords3.xml'
for event, elem in et.iterparse(bdata, events=('start', 'end')):
    if elem.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}r' and event == 'start':
       tmpdata = []
       for cidx, valueobj in enumerate(elem.getchildren()):
           tagname = valueobj.tag
           vattrib = valueobj.attrib.get('v')
           rdata = vattrib
           if tagname == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}x':
                try:
                  rdata = defdict[cidx][int(vattrib)]
                except:
                  logging.error('this it not should happen index cidx = {} vattrib = {} defaultidcts = {} tmpdata for the time = {} xml raw {}'.format(
                                                                                                                                                cidx, vattrib, defdict, tmpdata,
                                                                                                                                                et.tostring(list, encoding='utf8', method='xml')
                                                                                                                                                ))
           tmpdata.append(rdata)
       if tmpdata:
           dfdata.append(tmpdata)
       elem.clear()

fuel_by_uf = pd.DataFrame(dfdata)
fuel_by_uf.columns = ["combustivel", "ano", "regiao", "estado", "unidade", "jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez", "total"]

for column in ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez", "total"]:
    fuel_by_uf[column] = round(fuel_by_uf[column].astype(float),5)

diesel_by_fuel.to_csv("exame_raizen/output/diesel_by_fuel.csv", sep=";", encoding="Latin1", index="true")
