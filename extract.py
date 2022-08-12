import pandas as pd

class Extract:
    
    def extract_ot(path):
        ot=[]
        try:
            ot=pd.read_excel(path,sheet_name='R4_OT')
        except:
            print('No file R4_OT1')
        
        if len(ot)==0:
            try:
                ot=pd.read_excel(path,sheet_name='OT')
            except:
                print('No file R4_OT2')
        
        if len(ot)==0:
            try:
                ot=pd.read_excel(path,sheet_name='OT MB')
            except:
                print('No file R4_OT3')
        
        if len(ot)==0:
            try:
                ot=pd.read_excel(path,sheet_name='R04_T&D')
            except:
                print('No file R4_OT4')
        return ot

    def extract_hdd(path):
        hdd=[]
        try:
            hdd=pd.read_excel(path,sheet_name='R8_HDD')
        except:
            print('No file R8_HDD1')
        
        if len(hdd)==0:
            try:
                hdd=pd.read_excel(path,sheet_name='HDD')
            except:
                print('No file HDD2')
        
        if len(hdd)==0:
            try:
                hdd=pd.read_excel(path,sheet_name='R08_HDD')
            except:
                print('No file R4_OT3')
        
        
        return hdd    
    
    def extract_drt(path):
        drt=[]
        try:
            drt=pd.read_excel(path,sheet_name='R9_DRT')
        except:
            print('No file R9_DRT1')
        
        if len(drt)==0:
            try:
                drt=pd.read_excel(path,sheet_name='DRT')
            except:
                print('No file DRT2')
        
        if len(drt)==0:
            try:
                drt=pd.read_excel(path,sheet_name='R09_DRT')
            except:
                print('No file R9_DRT3')
        
        if len(drt)==0:
            try:
                drt=pd.read_excel(path,sheet_name='R09-DRT ')
            except:
                print('No file R9_DRT4')
        
        
        return drt    

    def extract_blo(path):
        blo=[]
        try:
            blo=pd.read_excel(path,sheet_name='R10_Blowing')
        except:
            print('No file R10_Blowing1')
        
        if len(blo)==0:
            try:
                blo=pd.read_excel(path,sheet_name='BLOWING')
            except:
                print('No file BLOWING2')
        
        if len(blo)==0:
            try:
                blo=pd.read_excel(path,sheet_name='Blowing')
            except:
                print('No file Blowing3')
        if len(blo)==0:
            try:
                blo=pd.read_excel(path,sheet_name='R10_BLOWING')
            except:
                print('No file Blowing4')
        
        
        return blo        
