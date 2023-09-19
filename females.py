# -*- coding: utf-8 -*-
"""
Get Female Height standards
"""

import pandas as pd
import numpy as np


def get_females(standard,heights):
    

    
    girlsthird = pd.Series(standard["Girls"]["3rd Percentile"])
    girlsmed = pd.Series(standard["Girls"]["Median (50th)"])
    girls97th = pd.Series(standard["Girls"]["97th percentile"])
    
    heightsfemale = heights[heights["Sex"]=="Female"]
    heightsfemale = pd.merge(heightsfemale,girlsthird,how='inner',left_on="Age",right_on=girlsthird.index)
    heightsfemale = pd.merge(heightsfemale,girlsmed,how='inner',left_on="Age",right_on=girlsmed.index)
    heightsfemale = pd.merge(heightsfemale,girls97th,how='inner',left_on="Age",right_on=girls97th.index)
    heightsfemale.sort_values(by=["Patient","VisitDate"],inplace=True)
    
    needed = pd.Series(heightsfemale.groupby("Patient")["CurrentHeight","3rd Percentile"].first().sum(axis=1), name = "HNAT")
    heightsfemale = pd.merge(heightsfemale,needed,how='inner',left_on="Patient",right_on=needed.index)
    
    heightsfemale['WeeklyGrowthRate'] = (
    (heightsfemale.groupby('Patient')['CurrentHeight'].transform('last') - heightsfemale.groupby('Patient')['CurrentHeight'].transform('first')) /
    ((heightsfemale.groupby('Patient')['VisitDate'].transform('last') - heightsfemale.groupby('Patient')['VisitDate'].transform('first')).dt.days / 7)
)

    heightsfemale["EstimatedWk52Date"] = pd.to_datetime(heightsfemale["EstimatedWk52Date"])
    weeksleft=pd.Series((heightsfemale.groupby('Patient')['EstimatedWk52Date'].transform('last') - heightsfemale.groupby('Patient')['VisitDate'].transform('last')).dt.days / 7)
    heightsfemale["ProjectedGrowth"] = ((weeksleft*heightsfemale["WeeklyGrowthRate"]) + heightsfemale.groupby('Patient')['CurrentHeight'].transform('last'))
    
    heightsfemale = heightsfemale.set_index(["Patient","VisitNum"])
    
    heightsfemale = heightsfemale.round(3)
    return heightsfemale