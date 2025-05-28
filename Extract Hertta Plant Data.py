# -*- coding: utf-8 -*-
"""
Created on Mon May 26 07:33:31 2025

@author: hajarals
"""

import numpy as np
import matplotlib.pyplot as plt
from scipy import optimize
import os
import pandas as pd 
##########################################################################################################################################################################################################
##########################################################################################################################################################################################################
path = str(input('Insert the path to the file:'))
num_point = int(input('insert number of measurement points:'))
Area = [] 
for i in range(num_point): 
    points = str(input(f'Insert name of measurement point {[i+1]}:'))
    Area.append(points) 
index = [] 
for i in range(65, 91): #there is length from 0, 255 including alphabatic and non alphabatic characters like / * : ...etc A:Z Have the index 65:91
    index.append(chr(i)) 
index.append('AA'); index.append('AB'); index.append('AC'); index.append('AD')
index.append('AE'); index.append('AF'); index.append('AG'); index.append('AH')
index.append('AI'); index.append('AJ'); index.append('AK'); index.append('AL')
index.append('AM'); index.append('AN'); index.append('AO');
data = pd.read_excel(path, index_col = None, sheet_name = 'Report results', skiprows=0, nrows=14645, usecols='A:AO')
nimi = data.iloc[:, 3].values.flatten() 
######################################################
Number = np.zeros(len(Area))
In = [[] for _ in range(len(Area))] 
rows = list(data.index)
for k in range(len(Area)):
    for i in range(len(nimi)):
        if Area[k] == nimi[i]: 
            Number[k]+=1
            In[k].append(i) 
######################################################
h = 0
skip = [] 
for i in range(len(In)):
    skip.append(In[i][0]) 
##########################################################################################################################################################################################################
##########################################################################################################################################################################################################
column = list(data.head(0)) 
col = ['Nimi', 'Pvm', 'Luokka', 'Lahko', 'Suku', 'Laji', 'Biomassa (µg/l)', 'Toksisuus', 'Haitallinen sinilevä', 'Flagellaatti', 'Nanoplankton', 'Kokoluokan kuvaus', 'Kpl/l']
##########################################################################################################################################################################################################
P = [[] for _ in range(len(Area))] 
for w1 in range(len(Area)):  
    mode1 = 'w'  
    path1 = str(input(f'Insert the path where to store the file:{[w1+1]}'))
    with pd.ExcelWriter(path1+'%s.xlsx'%Area[w1], engine='openpyxl', mode=mode1) as writer:
        VV = [[] for _ in range(len(col))] 
        for f in range(len(index)): 
            data1 = pd.read_excel(path, index_col = None, sheet_name = 'Report results', skiprows=skip[w1], nrows=int(Number[w1]), usecols=index[f])
            for j in range(len(col)):
                if col[j] == column[f]: 
                    VV[j].append(data1.iloc[:, 0].values.flatten())
                    P[w1].append(VV[j]) 
        place1, date1, class1, order1, family1, species1, value1, toxicity1, harm1, flag1, plank1, size1, kpl1 = [VV[i][0] for i in range(len(col))]
########################################################################################################################################################################################################
        dic = {'place': [place1[j] for j in range(len(place1))],
                'date': [date1[j] for j in range(len(date1))],
                'class': [class1[j] for j in range(len(class1))],
                'order': [order1[j] for j in range(len(order1))], 
                'family': [family1[j] for j in range(len(family1))],       
                'species': [species1[j] for j in range(len(species1))], 
                'value': [value1[j]/1000 for j in range(len(value1))],
                'Toxicity': [toxicity1[j] for j in range(len(toxicity1))],
                'Harmful': [harm1[j] for j in range(len(harm1))],
                'Flagellate': [flag1[j] for j in range(len(flag1))],
                'Nanoplankton': [plank1[j] for j in range(len(plank1))],
                'Size': [size1[j] for j in range(len(size1))],
                'kpl': [kpl1[j] for j in range(len(kpl1))]}                                                                                                                                     
        if '%s'%Area[w1] not in writer.book.sheetnames:
            sheet_name2 = '%s'%Area[w1] 
            Data3 = pd.DataFrame.from_dict(dic, orient='index').transpose()
            Data3.to_excel(writer, sheet_name=sheet_name2, startrow=0, startcol=0, index=False)     
########################################################################################################################################################################################################
########################################################################################################################################################################################################                                                 













