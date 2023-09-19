# -*- coding: utf-8 -*-


from myformatter import MyFormatter
import pandas as pd
from males import get_males
from females import get_females
from heightstandperc import get_standards
from getheight import get_heights
from datetime import date
import os

def main():
    standard = get_standards()
    height = get_heights()
    heightsfemale = get_females(standard,height)
    heightsmale = get_males(standard,height)
    
    last_row = len(heightsfemale) + 1
    last_rowm = len(heightsmale) + 1
    with pd.ExcelWriter("HeightMonitoring{:%d-%b}.xlsx".format(date.today())) as writer:
        workbook = writer.book
        danger_format = workbook.add_format({'border': 1, 'bg_color': '#ff3333'})
        clear_format = workbook.add_format({'border': 1, 'bg_color': '#00ff00'})
        date_format = workbook.add_format({'num_format': 'dd-mmm-yyyy','border': 1})
        bold = workbook.add_format({'bold': True, 'font_color': 'black'})
        bold.set_align("top")
        red = workbook.add_format({'bold': True, 'font_color': '#ff0000'})
        red.set_align('vcenter')
        red.set_text_wrap()
        green = workbook.add_format({'bold': True, 'font_color': '#248f24'})
        green.set_align('vcenter')
        green.set_text_wrap()
        wrap = workbook.add_format({'bold': False, 'font_color': 'black'})
        wrap.set_align('vcenter')
        wrap.set_text_wrap()
        
        heightsfemale.to_excel(writer, sheet_name="Females")
        
        worksheet = writer.sheets["Females"]
        
        # If Height Needed to be in 3rd percentile > Projected Height at final visit
        formula_danger = f'=L2>N2'
        formula_clear = f'=G2>L2'
        
        worksheet.conditional_format(f'A2:A{last_row}', {'type': 'formula',
                                                        'criteria': formula_danger,
                                                        'format': danger_format})
        worksheet.conditional_format(f'A2:A{last_row}', {'type': 'formula',
                                                    'criteria': formula_clear,
                                                    'format': clear_format})
        worksheet.conditional_format(f'G2:G{last_row}', {'type': 'formula',
                                                    'criteria': formula_clear,
                                                    'format': clear_format})
        worksheet.conditional_format(f'G2:G{last_row}', {'type': 'formula',
                                                    'criteria': formula_danger,
                                                    'format': danger_format})
        worksheet.conditional_format(f'F2:F{last_row}', {'type': 'no_errors','format': date_format})
        worksheet.conditional_format(f'H2:H{last_row}', {'type': 'no_errors','format': date_format})
        formatter = MyFormatter(heightsfemale,workbook,worksheet)
        formatter.spacing()
        formatter.color_columns(color="#b3e0ff")
        formatter.add_borders()
        worksheet.freeze_panes(1,1)
        # Males
        heightsmale.to_excel(writer, sheet_name="Males")
        
        worksheet = writer.sheets["Males"]
        worksheet.conditional_format(f'A2:A{last_row}', {'type': 'formula',
                                                        'criteria': formula_danger,
                                                        'format': danger_format})
        worksheet.conditional_format(f'A2:A{last_row}', {'type': 'formula',
                                                    'criteria': formula_clear,
                                                    'format': clear_format})
        worksheet.conditional_format(f'G2:G{last_row}', {'type': 'formula',
                                                    'criteria': formula_clear,
                                                    'format': clear_format})
        worksheet.conditional_format(f'G2:G{last_row}', {'type': 'formula',
                                                    'criteria': formula_danger,
                                                    'format': danger_format})
        worksheet.conditional_format(f'F2:F{last_rowm}', {'type': 'no_errors','format': date_format})
        worksheet.conditional_format(f'H2:H{last_rowm}', {'type': 'no_errors','format': date_format})
        formatter = MyFormatter(heightsmale,workbook,worksheet)
        formatter.spacing()
        formatter.color_columns(color="#b3e0ff")
        formatter.add_borders()
        worksheet.freeze_panes(1,1)
        
        pd.DataFrame().to_excel(writer,sheet_name="Definitions", index = False)
        worksheet=writer.sheets["Definitions"]
        worksheet.set_column(0, 0, 20)
        worksheet.set_column(1, 1, 35)
        worksheet.write(0, 0, 'HNAT:',bold)
        worksheet.write(0, 1, 'Final Week 52 Height needed for a patient to be in the 3rd Percentile.\nCalculated as BaselineWeight + 3rd Percentile Growth Rate for that specific age and sex cohort.',wrap)
        worksheet.write(1,0,"WeeklyGrowthRate:",bold)
        worksheet.write(1,1,"Most Recent Visit Height - Baseline Height/Time between those visits in weeks",wrap)
        worksheet.write(2,0,"ProjectedGrowth:",bold)
        worksheet.write(2,1,"(WeeklyGrowthRate * Number of weeks from most recent vist date to EstimatedWk52Date)\n + Most recent CurrentHeight",wrap)
        worksheet.write(4,1,"Patients Highlighted in Red are not not expected to be above the 3rd percentile based on current GR",red)
        worksheet.write(6,1,"Any CurrentHeight highlighted in Greeen is already above the 3rd percentile",green)
        worksheet.write(8,1,"Percentile Standards are taken from the Table of Yearly growth velocity by age and gender(abstracted  from Tanner and  Davis. J Pediatr. 1985;107(3):317-329)",wrap)
    os.system("start EXCEL.EXE HeightMonitoring{:%d-%b}.xlsx".format(date.today()))

if __name__ == "__main__":
    main()
    