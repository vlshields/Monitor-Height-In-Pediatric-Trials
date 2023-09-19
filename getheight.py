# -*- coding: utf-8 -*-
"""
This script pulls the line listing.
If your trial lasts longer than 52 weeks,
the codebase will have to be adjusted.
"""

from quickhelper import QuickHelper
import pandas as pd


def get_heights():
    
    qh = QuickHelper("heightpull")
    heights = qh.data
    
    # We are only looking at patients currently enrolled and who have not had week 52
    heights = heights[heights["Patient Disposition"] == "Enrolled"]
    heights = heights[~heights.groupby("Patient")["VisitNum"].transform(lambda x: x.eq("52").any())]
    
    # Keep only the necessary columns
    heights = heights[['Patient', 'Interval Value', 'StudyDay', 'Patient Disposition',
           'Age', 'Sex', 'Race', 'EnrollDate', 'VisitNum', 'VisitDate',
           'WeeksUntilWeek52Expected', 'EstimatedWk52Date', 'CurrentHeight']]
    
    # Sort the values by Patient first then by visit date
    heights["VisitDate"] = pd.to_datetime(heights["VisitDate"])
    heights.sort_values(by=["Patient","VisitDate"],inplace=True)
    
    # We are also excluding prebaseline visits
    heights = heights[["Patient",'Age', 
                       'Sex','VisitNum','Patient Disposition', 
                       'VisitDate',"CurrentHeight",
                       'EstimatedWk52Date']][heights["VisitNum"] != "Pre-BL"]
    return heights
