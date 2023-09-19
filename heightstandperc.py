# -*- coding: utf-8 -*-
"""
Script cleans the the height percentiles
"""

import pandas as pd
import numpy as np

def get_standards():
    
    standards = pd.read_excel("heightstandards.xlsx","Sheet1")
    standards = standards.dropna(axis=1,how='all')
    
    
    standard = pd.DataFrame({"Age":list(standards.iloc[1:,0]),
                            "Boys_Median (50th)":list(standards.iloc[1:,1]),
                            "Boys_3rd Percentile":list(standards.iloc[1:,2]),
                            "Boys_97th percentile":list(standards.iloc[1:,3]),
                            "Girls_Median (50th)":list(standards.iloc[1:,4]),
                            "Girls_3rd Percentile":list(standards.iloc[1:,5]),                  
                            "Girls_97th percentile":list(standards.iloc[1:,6])})
    
    # Set Age as index and create hierarchical columns
    standard = standard.set_index("Age")
    standard.columns = pd.MultiIndex.from_tuples([tuple(c.split("_")) for c in standard.columns])
    
    # Keep only whole numbers because we are using age at screening
    for index in standard.index:
        if '.5' in str(index):
            standard.drop(index,inplace=True)
            
    
    # Ensure data type as integers
    standard.index = standard.index.astype('int')
    
    return standard





