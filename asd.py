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
from tkinter import messagebox

file2=""
file1=""
popupRoot = Tk()

popupRoot.title('Generate BoQ')



def boq_format():
    global file2
    t=filedialog.askopenfilename(initialdir="C:/")
    file2=file2+t
    print(t,file2)


def mandal_folder(): 
    global file1   
    # file_path = filedialog.askopenfilename()
    t=filedialog.askdirectory()
    file1=file1+t
    print(t,file1)

def boq12():
    ##Getting BoQ formatted version to write directly
    xfile = openpyxl.load_workbook(file2)
    #sheet = xfile['GP-Wise BOQ']
    try:
        sheet1=xfile['GP-Wise BOQ']
    except:
        messagebox.showerror("Extract Error", "BoQ file should have Sheet name 'GP-Wise BOQ'")
    #print(len(Sheet1['H']))
    #print(Sheet1['F3'].value)

    mandal_file=file1+'/'
    print("Grabbed mandal location")
    # print(mandal_file)
    no=0
    os.chdir(mandal_file)
    for info in os.listdir():

        j=mandal_file+info
        print(j)

        blo=Extract.extract_blo(j)
        if len(blo)!=0:
            (blo,joint_closer)=Create_table.create_blo(blo)

        #blo.to_csv('blwoing.csv')
        #print("pas printing")

        drt=Extract.extract_drt(j)
        if len(drt)!=0:
            drt=Create_table.create_drt(drt)

        #print("empty")
        #print(drt)

        #drt.to_csv('drt12.csv')
        ot=Extract.extract_ot(j)
        if len(ot)!=0:
            ot=Create_table.create_ot(ot)   

        hdd=Extract.extract_hdd(j)
        if len(hdd)!=0:
            hdd=Create_table.create_hdd(hdd)      
        #print(joint_closer)
        s=0
        fin=0
        temp=0

        try:
            #print(range(len(drt)))
            #print(drt.iloc[0])
            sum=0
            #drt.to_csv('drt133.csv')
            drt=drt.reset_index(drop=True)
            for i in range(len(drt)):
                ch1=drt.loc[i,'ch_from']
                ch2=drt.loc[i,'ch_to']
                if i!=(len(drt)-1):
                    ch3=drt.loc[i+1,'ch_from']
                    if(ch3!=ch2):
                        sum+=ch3-ch2
            print("sum after chaining")
            print(sum)

            if np.isnan(sum):
                sum=0

            for i in range(len(drt)):
                ch1=drt.loc[i,'Duct_miss_ch_from']
                ch2=drt.loc[i,'Duct_miss_ch_to']
                if i!=(len(drt)-1):
                    ch3=drt.loc[i+1,'Duct_miss_ch_from']

                temp=(drt.loc[i,'Duct_miss_ch_Length'])

                if ch3==ch2:
                    s+=temp  

                elif ch3!=ch2:
                    #print("got it")
                    s=np.nansum([s+temp])
                    #print(temp,"s is",s,"done")
                    #print(s)
                    if(s>=50):
                        s-=50
                        fin+=s
                        s=0
                    elif s<50:
                        s=0
                #print(temp," ",s," ",fin," ",ch1," ",ch2," ",ch3)          
        except:
            print("Something Wrong in missing calculation")
        print("Sum after missing")
        sum+=fin
        print(fin)                

        try:
            blo=blo.reset_index(drop=True)
            start_drt=drt['ch_from'].iloc[0]
            sum+=(start_drt-0)
            print("Sum before starting 0")
            print(sum)
            last_blow=blo['Chainage_To'].iloc[-1]
            if isnan(last_blow):
                last_blow=blo['Chainage_To'].iloc[-2]
            print(last_blow)
            last_drt=drt['ch_to'].iloc[-1]
            print(last_drt)
            print("Final sum")
            sum+=last_blow-last_drt

            print(sum)    


            #ot.to_csv('ot12.csv')    

            #print(sheet[chr(ord('F')+no)+'9'])
        except:
            print('Something wrong in blowing gap calculation or drt file')    

        var='F'
        var1=chr(ord(var)+1)
        var=chr(ord(var)+2*no)
        try:
            sheet1[var+'3']=info
            sheet1[var1+'4']='STATE'
            sheet1[var+'4']='BBNL'
            sheet1[var+'11']=(blo.loc[blo['size_of_ofc']=='288F','Total_cable_length'].sum())/1000
            sheet1[var+'14']=(blo.loc[blo['size_of_ofc']=='144F','Total_cable_length'].sum())/1000
            sheet1[var+'17']=(blo.loc[blo['size_of_ofc']=='96F','Total_cable_length'].sum())/1000
            sheet1[var+'18']=(blo.loc[blo['size_of_ofc']=='48F','Total_cable_length'].sum())/1000
        except:
            print("No such Blowing files..")  
        try:
            a=(drt['Duct_miss_ch_Length']==0)
            b=~(drt['Duct_miss_ch_Length'].astype(str).str.isdigit())
            y=a+b
            #print(a+b)
            c=(drt['Duct_dam_punct_loc_Length']==0)
            d=~(drt['Duct_dam_punct_loc_Length'].astype(str).str.isdigit())
            #print(c+d)
            z=c+d
            e=y*z
            sheet1[var+'99']=drt.loc[e,'Length'].sum()/1000
            sheet1[var+'22']=sum/1000
        except:
            print("No such DRT files..")  
        try:
            sheet1[var+'23']=(ot['Protect_Dwc'].sum())/1000
            sheet1[var+'24']=(ot['Protect_Gi'].sum())/1000
            sheet1[var+'25']=(ot['Rcc_Chamber'].sum())
            sheet1[var+'26']=(ot['Rcc_Marker'].sum())

        except:
            print("No such OT files..")    
        xfile.save(filename="BOQ122.xlsx")
        ##print(duct_laid,dwc_pipe,gi_pipe,route_indicators,joint_chambers)
        no+=1
    popupRoot.destroy()


popuplabel = Label(popupRoot, text = 'Please select BoQ Format First',font = ("Times New Roman", 10)).grid(row=0,column=1)
popupButton = Button(popupRoot, text = 'Upload BoQ Format', font = ("Times New Roman", 10), command = boq_format,width = 20).grid(row=0,column=2)
popuplabel = Label(popupRoot, font = ("Times New Roman", 10),text = 'Upload Mandal Folder').grid(row=1,column=1)
popupButton = Button(popupRoot, text = 'Upload Mandal Folder', font = ("Times New Roman", 10), command = mandal_folder,width = 20).grid(row=1,column=2)
popupButton = Button(popupRoot, text = 'Run BoQ Script', font = ("Times New Roman", 11), command = boq12,width = 15).grid(row=3,column=1)
popupRoot.geometry('350x100')
#print("1")

#print("printing")
#print(file2)
#print(file1)
#print("1")

#popupRoot.destroy()
mainloop()
