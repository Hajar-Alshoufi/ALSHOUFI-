# -*- coding: utf-8 -*-
"""
Created on Fri May 16 16:39:30 2025

@author: hajarals
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import os
###########################################################################################################################################################
###########################################################################################################################################################
path = str(input('Inser the path to the excel file:')) #without ''
rangee = str(input('Insert the range of data:'))
num = int(input('Insert the number of sheets in excel file:'))
co = int(input('Insert the number of nutrient codes:'))
nume = int(input('Insert the number of meaurement places on the map:'))
Name = [] 
Number = [] 
for i in range(num):
    Nam = str(input(f'Insert name of the excel sheet{[i+1]}:'))
    rows = int(input(f'Insert the total number of rows in sheet{[i+1]}:'))
    Name.append(Nam)
    Number.append(rows) 
Code = [] 
for i in range(co):
    CO = str(input(f'Insert nutrient code{[i+1]}:'))
    Code.append(CO)
Area = [] 
for i in range(nume): 
    plac = str(input(f'Insert the meaurement place on the map{[i+1]}:'))
    Area.append(plac)
epi_depth = str(input('Insert the epilimnion depth:'))
hypo_depth = str(input('Insert the hypolimnion depth:'))
epi_depth = '%s,0'%epi_depth
hypo_depth = '%s,0'%hypo_depth
column = ['Paikan nimi', 'Paikan ID-numero', 'Paikan ETRS-koord itä', 'Paikan ETRS-koord pohj',	'Paikan syvyys (m)', 'Näytteenottoaika', 'Näytesyvyys',	'Määrityskoodi', 'Suure(suomeksi)', 'Lippu', 'Alkuperäinen arvo', 'Yksikkö', 'Tulos']
col = ['Paikan nimi', 'Näytteenottoaika', 'Näytesyvyys', 'Suure(suomeksi)', 'Tulos']
index = [chr(i) for i in range(ord(rangee[0]), ord(rangee[-1])+1)]
########################################################################################################################################################################################################################################
########################################################################################################################################################################################################################################
for w1 in range(len(Area)):
    mode1 = 'w'
    DEC = [{} for i in range(len(Code))]
    path1 = str(input(f'Insert path where to store file{[w1+1]}:')) 
    with pd.ExcelWriter(path1%Area[w1], engine='openpyxl', mode=mode1) as writer:
        for c in range(len(Name)):
            VV = [[] for _ in range(len(col))]
            for f in range(len(index)): 
                data1 = pd.read_excel(path, index_col = None, sheet_name = Name[c], skiprows=0, nrows=Number[c], usecols=index[f])
                data1.fillna('0,0', inplace=True)
                for j in range(len(col)):
                    if col[j] == column[f]:
                        VV[j].append(data1.iloc[:, 0].values.flatten())
            place1, date1, depth1, code1, value1 = [VV[i][0] for i in range(len(col))]
            for b in range(len(Code)):
                epi = [[] for _ in range(len(Code))]; hypo = [[] for _ in range(len(Code))]; 
                date_epi = [[] for _ in range(len(Code))]; date_hypo = [[] for _ in range(len(Code))] 
                for i, j, k, l, m in zip(date1, depth1, value1, place1, code1): 
                    if Area[w1] == l and Code[b] == m:
                        if j == '0,0-2,0':
                            if m == 'Klorofylli-a':
                                epi[b].append(k)  
                                date_epi[b].append(i) 
                        elif j == '0,0':
                            if m == 'Näkösyvyys':
                                epi[b].append(k)
                                date_epi[b].append(i)
                        elif j == epi_depth:
                            if m == 'Kokonaistyppi, suodattamaton' or m == 'Kokonaisfosfori, suodattamaton' or m == 'Nitriitti-nitraatti typpenä, suodattamaton' or m == 'Ammonium typpenä, suodattamaton':
                                epi[b].append(k/1000)  
                                date_epi[b].append(i)
                            else:
                                epi[b].append(k)  
                                date_epi[b].append(i)
                        elif j == hypo_depth:
                            if m == 'Kokonaistyppi, suodattamaton' or m == 'Kokonaisfosfori, suodattamaton' or m == 'Nitriitti-nitraatti typpenä, suodattamaton' or m == 'Ammonium typpenä, suodattamaton':
                                hypo[b].append(k/1000)  
                                date_hypo[b].append(i)
                            else:
                                hypo[b].append(k)  
                                date_hypo[b].append(i) 
########################################################################################################################################################################################################################################
########################################################################################################################################################################################################################################                                
                if epi[b] != [] or hypo[b] != [] or date_epi[b] != [] or date_hypo[b] != []:  
                    dic = {'date_epi': [date_epi[b][j] for j in range(len(date_epi[b]))],
                            f'{Code[b]}_epi': [epi[b][j] for j in range(len(epi[b]))],
                            'date_hypo': [date_hypo[b][j] for j in range(len(date_hypo[b]))], 
                            f'{Code[b]}_hypo': [hypo[b][j] for j in range(len(hypo[b]))]}
                    DEC[b] = dic 
                    for d1 in DEC[b]: 
                        if DEC[b][d1] != []:
                            DEC[b] = dic 
                            if '%s'%Code[b] not in writer.book.sheetnames:
                                sheet_name2 = '%s'%Code[b]
                                if len(sheet_name2) > 31:
                                    sheet_name2 = '%s'%Code[b][:15]
                                else:
                                    sheet_name2 = '%s'%Code[b] 
                                    Data3 = pd.DataFrame.from_dict(DEC[b], orient='index').transpose()
                                    Data3.to_excel(writer, sheet_name=sheet_name2, startrow=0, startcol=0, index=False)    
                elif epi[b] == [] or hypo[b] == [] or date_epi[b] == [] or date_hypo[b] == []:
                    dic5 = {'date': [], f'{Code[b]}': []}
                    DEC[b] = dic5
                    if '%s'%Code[b] not in writer.book.sheetnames:
                        sheet_name5 = '%s'%Code[b]
                        if len(sheet_name5) > 31:
                            sheet_name5 = '%s'%Code[b][:15]
                        else:
                            sheet_name5 = '%s'%Code[b]
                            Data4 = pd.DataFrame.from_dict(DEC[b], orient='index').transpose()
                            Data4.to_excel(writer, sheet_name=sheet_name5, startrow=0, startcol=0, index=False) 
###################################################################################################################################################################   
###################################################################################################################################################################          
