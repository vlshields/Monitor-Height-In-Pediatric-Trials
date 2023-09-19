# -*- coding: utf-8 -*-

"""Quick Excel Formatter. Custom excel format that adds
spacing and color of your chosing"""

import pandas as pd
import glob
import os
from pathlib import Path
import win32com.client


class MyFormatter:
    def __init__(self, data, workbook, worksheet):
        self.data = data
        self.workbook = workbook
        self.worksheet = worksheet
    
    def spacing(self):
        """
        Function to add spacing to Excel columns based on the maximum length of the
        data in each column.
        """
        if isinstance(self.data, pd.DataFrame):
            if self.data.index.nlevels > 1:
                self.data = self.data.reset_index()
        elif isinstance(self.data, pd.core.groupby.DataFrameGroupBy):
            self.data = self.data.apply(lambda x: x.reset_index(drop=True))
            self.data = self.data.reset_index()
        else:
            raise TypeError("The data must be a pd.DataFrame or a pd.core.groupby.DataFrameGroupBy.")

        for i, col in enumerate(self.data.columns):
            max_width = max(self.data[col].astype(str).str.len().max(), len(col))
            self.worksheet.set_column(i, i, max_width + 1)

    def color_columns(self, grey_bottom=False, color='#ffb3b3'):
        """
        Function to color columns in Excel worksheet.
        """
        if isinstance(self.data, pd.DataFrame):
            if self.data.index.nlevels > 1:
                self.data = self.data.reset_index()
        elif isinstance(self.data, pd.core.groupby.DataFrameGroupBy):
            self.data = self.data.apply(lambda x: x.reset_index(drop=True))
            self.data = self.data.reset_index()
        else:
            raise TypeError("The data must be a pd.DataFrame or a pd.core.groupby.DataFrameGroupBy.")

        if grey_bottom:
            last_row = self.data.shape[0]
            last_column = self.data.shape[1]
            last_row_format = self.workbook.add_format({'bg_color': '#d9d9d9', 'bold': True})
            self.worksheet.conditional_format(last_row, 0, last_row, last_column, {'type': 'no_blanks', 'format': last_row_format})

        column_format = self.workbook.add_format({'bg_color': color})
        self.worksheet.conditional_format(0, 0, 0, len(self.data.columns), {'type': 'no_blanks', 'format': column_format})

    def add_borders(self):
        """
        Function to add borders to cells in Excel worksheet.
        """
        if isinstance(self.data, pd.DataFrame):
            if self.data.index.nlevels > 1:
                self.data = self.data.reset_index()
        elif isinstance(self.data, pd.core.groupby.DataFrameGroupBy):
            self.data = self.data.apply(lambda x: x.reset_index(drop=True))
            self.data = self.data.reset_index()
        else:
            raise TypeError("The data must be a pd.DataFrame or a pd.core.groupby.DataFrameGroupBy.")

        border_format = self.workbook.add_format({'border': 1, 'border_color': 'black'})
        border_format.set_locked(False)
        self.worksheet.conditional_format(1, 0, len(self.data), self.data.shape[1]-1, {'type': 'no_errors', 'format': border_format})
        
    def panes(self):
        self.worksheet.freeze_panes(1,1)

    



class ReNamer:
    def __init__(self,filename):
        
        self.filename = filename
        self.type = self.file_type()
        
    def file_type(self):

        if self.filename.lower().endswith('.xlsx'):
            return "Excel"
        elif self.filename.lower().endswith('.csv'):
            return "CSV"
        elif self.filename.lower().endswith('.png'):
            return "PNG"
        else:
            raise TypeError("File extension not recognized. Takes .xlsx,.csv,.png") 
        
    def move_to_cwd(self):
        
        if self.type == "CSV":
            csv_files = glob.glob(r"C:\Users\vshields\Downloads\*.csv")
            csv_files_sorted = sorted(csv_files, key=os.path.getmtime, reverse=True)
            
            try:
                csv = pd.read_csv(csv_files_sorted[0])
            except UnicodeDecodeError as e:
                print(e)
                csv = pd.read_csv(csv_files_sorted[0],encoding = 'unicode_escape', engine ='python')
            
                
            csv.to_csv(self.filename, index=False)
            
        
        elif self.type == "Excel":
            
            files = glob.glob(r"C:\Users\vshields\Downloads\*.xlsx")
            file = sorted(files, key=os.path.getmtime, reverse=True)[0]
            file = pd.read_excel(file)
        
            file.to_excel(str(Path.cwd()) + '\\'+self.filename, index=False)
            
        else:
            raise TypeError("Method only takes csv or xlsx")
            
    def convert_xls_to_xlsx(self):
        
        xls_files = glob.glob(r"C:\Users\vshields\Downloads\*.xls")
        xls_file = sorted(xls_files, key=os.path.getmtime, reverse=True)[0]
        
        # create an instance of Excel
        excel = win32com.client.Dispatch("Excel.Application")
        
        # open the XLS file
        workbook = excel.Workbooks.Open(xls_file)
        
        # save the XLS file as XLSX
        workbook.SaveAs(str(Path.cwd()) + '\\'+self.filename, FileFormat=51)
        
        # close the workbook
        workbook.Close()
        
        # quit Excel
        excel.Quit()


       
    
    def read_png(self):
        
        png_files = glob.glob(r"C:\Users\vshields\Desktop\Incyte 305\*.png")
        png_files_sorted = sorted(png_files, key=os.path.getmtime, reverse=True)
        
        return png_files_sorted[0]
    

                
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
    