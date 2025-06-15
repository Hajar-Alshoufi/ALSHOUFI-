# -*- coding: utf-8 -*-
"""
Created on Wed Jan  8 13:25:30 2025
@author: Dr. Hajar Alshoufi 

This code extract hourly wind data and store them into 
excel sheets where you can later organize them for each year separately
"""

import numpy as np
import matplotlib.pyplot as plt
from scipy import optimize
import os
import pandas as pd 
import sys
############################################################################################################################################################################################################
path = str(input('Insert the path where the file is stored:'))
sheet = str(input('Insert the name of excel sheet:'))
num = int(input('Insert the number of rows in the Excel sheet:'))
rangee = str(input('Insert the range in excel sheet:')) 
ye1 = int(input('Insert the starting year: '))
ye2 = int(input('Insert the ending year: '))
mon1 = int(input('Insert the starting month: '))
mon2 = int(input('Insert the ending month: '))
dis1 = int(input('Insert the starting day: '))
dis2 = int(input('Insert the ending day: '))
vuosi = np.arange(ye1, ye2+1, 1)
#######################################################
################Read the data from the Excel Sheet##############
data =  pd.read_excel(path, index_col = None, sheet_name='Kuopio Savilahti', skiprows=0, nrows= 104440, usecols="B:G")  
index = [chr(i) for i in range(ord(rangee[0]), ord(rangee[-1])+1)]
data =  pd.read_excel(path, index_col = None, sheet_name=sheet, skiprows=0, nrows=num, usecols=rangee)  
data.fillna('', inplace=True) 
head = list(data.head()) 
column = ['Vuosi', 'Kuukausi', 'Päivä', 'Keskituulen nopeus [m/s]']
VV = [[] for i in range(len(column))]
for f in range(len(index)): 
    data =  pd.read_excel(path, index_col = None, sheet_name=sheet, skiprows=0, nrows=num, usecols=index[f])  
    data.fillna('', inplace=True)
    for j in range(len(column)):
        if column[j] == head[f]:
            VV[j].append(data.iloc[:, 0].values.flatten())
year, month, day, wind = [VV[i][0] for i in range(len(column))]     
#####################################  
YY = [[]*i for i in range(len(vuosi))]
z = [4, 6, 9, 11]
o = [1, 3, 5, 7, 8, 10, 12]
V1 = ye1; 
Ve = ye2+1; 
q = 0; 
m = dis1; 
G = []
while V1 < Ve and q < len(vuosi):
    s1 = 13     #month counter
    kuukausi = np.arange(mon1, s1, 1) 
    k1 = mon1; L1 = 32; a = 0
    WW = [[]*i for i in range(s1)]
    while k1 < s1 and a < len(kuukausi):
        if k1 == mon2 and vuosi[q] == ye2 and mon2 == 2:
            s1 = mon2
            if dis2 == 1:
                L1 = dis2 + 4
            elif dis2 == 2:
                L1 = dis2 + 3
            elif dis2 == 3:
                L1 = dis2 + 3
        if k1 == mon2 and vuosi[q] == ye2 and mon2 != 2:
            L1 = dis2+1
            s1 = mon2+1
        if m < 0:
            m = 1
        for r in o:
            if k1 == r:
                da1 = np.zeros(L1)
        for kr in z:
            if k1 == kr:
                da1 = np.zeros(L1-1)
        if k1 == 2 and vuosi[q]%4 == 0:
            da1 = np.zeros(L1-2)
        elif k1 == 2 and vuosi[q]%4 != 0:
            da1 = np.zeros(L1-3) 
        k = m    # day counter 
        while k < len(da1):
            count = 0
            for j in range(len(day)):
                if k1 == month[j] and k == day[j] and V1 == year[j]:
                    if wind[j] != 0: 
                        da1[k]+=wind[j]
                        count+=1  
                        if f'{day[j]:02d}-{k1:02d}-{vuosi[q]}' not in G:
                            G.append(f'{day[j]:02d}-{k1:02d}-{vuosi[q]}')
            if count != 0:
                avg = float(da1[k]/count)
                WW[a].append(avg)
            k+=1 
        if k1 == 12: 
            mon1 = 1
        a+=1
        k1+=1
        m-=dis1-1    
    YY[q].append(WW)
    V1+=1
    q+=1
###############################################################################
ii = 0
wind_all = []
while ii < len(YY):                     #loop through the years
    for q in range(len(YY[ii])):        #loop through the months
        for j in range(len(YY[ii][q])): #loop through the months again
              if YY[ii][q][j] != []: 
                for i in range(len(YY[ii][q][j])):
                    wind_all.append(YY[ii][q][j][i])
    ii+=1
###############################################################################    
a = {'Date': [G[k] for k in range(len(G))],
     'Wind m/s': [wind_all[i] for i in range(len(wind_all))]} #get the days
path1 = str(input('Insert the path where to store the file:'))
Data = pd.DataFrame.from_dict(a, orient='index').transpose() 
Data.to_excel(path1+'Wind Date.xlsx', sheet_name='Wind', startrow=0, startcol=0, index=False)                 
############################################################################################################################################################################################################
############################################################################################################################################################################################################
############################################################################################################################################################################################################
