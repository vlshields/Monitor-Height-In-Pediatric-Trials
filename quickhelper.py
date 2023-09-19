# -*- coding: utf-8 -*-
"""
Download your line listing
and run this program. 
Change file extension if you prefer
"""
import pandas as pd
from glob import glob
import numpy as np
from datetime import date
import os

class QuickHelper:
    
    def __init__(self,datagrid):
        
        self.filename = datagrid + date.today().strftime("%d%b%Y") +'.csv'
        self.data = self.prep_data()
          
        
    def getfile(self,csv=True):
        
        recent = sorted(glob("C:\\Users\\vshields\\Downloads\\*.csv"),
                        key=os.path.getmtime,reverse=True)[0]
        os.rename(recent,self.filename)
    
    def prep_data(self):
        self.getfile()
        
        return pd.read_csv(self.filename,na_values = "-")
    
    
    

        

        
        
        
        
        
        
        
