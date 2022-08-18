from tkinter import *
from sys import exit
from tkinter import filedialog
import pandas as pd
import numpy as np
import os
from extract import Extract
from create_table1 import Create_table
import openpyxl
import matplotlib.pyplot as plt
from openpyxl_image_loader import SheetImageLoader
# import excel2img
import os
import tkinter
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
#Parth Pandey12:12 PM
pd.options.mode.chained_assignment = None

j=r'C:\Users\Aditya.gupta\Desktop\Mandal\WRG-ZAF-5706-M-01-GR01-13.xlsx'

drt=Extract.extract_drt(j)
if len(drt)!=0:
    drt=Create_table.create_drt(drt)
    drt.reset_index(drop=True,inplace=True)
    #print("empty")


a=(drt['Duct_miss_ch_Length']==0)
b=~(drt['Duct_miss_ch_Length'].astype(str).str.isdigit())
y=a+b
#print("Printing a+b",a+b)
#print("Dit length only miss",drt.loc[y,'Length'].sum()/1000)
c=(drt['Duct_dam_punct_loc_Length']==0)
d=~(drt['Duct_dam_punct_loc_Length'].astype(str).str.isdigit())
#print("Printing c+d",c+d)
z=c+d
#tim=(drt.loc[z,'Length'])
#print(tim.loc['Length'].sum())
#print(drt.loc[z,'Length'].sum())
#print("Dit length only dam",drt.loc[z,'Length'].sum()/10
e=y*z
drtmissseries=pd.Series(drt['Duct_miss_ch_Length'])
drtmissseriesbool=pd.Series(y)
drtdamseries=pd.Series(drt['Duct_dam_punct_loc_Length'])
drtdamseriesbool=pd.Series(z)
drtdammissseries=pd.Series(e)
drtrequirelen=pd.Series(drt['Length'])
drttype=pd.Series(type(drt['Length']))
tim={'DRT mis length':drtdammissseries,'Miss bool value':drtmissseriesbool,'DRT dam length':drtdamseries,'Dam bool value':drtdamseriesbool,'Dam+dis bool':drtdammissseries,'Required Length':drtrequirelen,'Req length type':drttype}
print(tim)
tim_df=pd.DataFrame(tim)
#tim=drt.loc[e,'Length']
print("tim dataframe is getting printed")
tim_df.to_csv('tim.csv')
#print(tim)
#print("Printing tim",tim)
#mum=0

print("Printing dit sum",drt.loc[e,'Length'].sum())

#print(drt.loc[e,'Length'].sum())