# -*- coding: utf-8 -*-
""""Get Male Height Standards"""

import pandas as pd
import numpy as np
from heightstandperc import get_standards
from getheight import get_heights

def get_males(standard,heights):
    

    
    boysthird = pd.Series(standard["Boys"]["3rd Percentile"])
    boysmed = pd.Series(standard["Boys"]["Median (50th)"])
    boys97th = pd.Series(standard["Boys"]["97th percentile"])
    

    
    heightsmale = heights[heights["Sex"]=="Male"]
    
    # Add nale height percentiles
    heightsmale = pd.merge(heightsmale,boysthird,how='inner',left_on="Age",right_on=boysthird.index)
    heightsmale = pd.merge(heightsmale,boysmed,how='inner',left_on="Age",right_on=boysmed.index)
    heightsmale = pd.merge(heightsmale,boys97th,how='inner',left_on="Age",right_on=boys97th.index)
    heightsmale.sort_values(by=["Patient","VisitDate"],inplace=True)
    
    # Calculate the height needed per patient on final visit for patient to be
    # above the third percentile
    needed = pd.Series(heightsmale.groupby("Patient")["CurrentHeight",
                                                      "3rd Percentile"].first().sum(axis=1),
                       name = "HNAT")
    heightsmale = pd.merge(heightsmale,needed,how='inner',left_on="Patient",right_on=needed.index)
    
    # Calculate latest measured height subtracted by their baseline hight / the time between those visits in days
    heightsmale['WeeklyGrowthRate'] = (
        (heightsmale.groupby('Patient')['CurrentHeight'].transform('last') - heightsmale.groupby('Patient')['CurrentHeight'].transform('first')) /
        ((heightsmale.groupby('Patient')['VisitDate'].transform('last') - heightsmale.groupby('Patient')['VisitDate'].transform('first')).dt.days / 7)
    )
    
    # Calculate projected growth
    heightsmale["EstimatedWk52Date"] = pd.to_datetime(heightsmale["EstimatedWk52Date"])
    weeksleft=pd.Series((heightsmale.groupby('Patient')['EstimatedWk52Date'].transform('last') - heightsmale.groupby('Patient')['VisitDate'].transform('last')).dt.days / 7)
    heightsmale["ProjectedGrowth"] = ((weeksleft*heightsmale["WeeklyGrowthRate"]) + heightsmale.groupby('Patient')['CurrentHeight'].transform('last')) 
    
    heightsmale = heightsmale.set_index(["Patient","VisitNum"])
    
    heightsmale = heightsmale.round(3)
    return heightsmale