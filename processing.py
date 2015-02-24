__author__ = 'abuteau'

import os
import sys
from xlwings import Workbook, Range, Sheet
import pandas as pd
from pandas import ExcelWriter



def cleanData(df):

    df = df[pd.notnull(df['V14'])]   # Removing empty entries from V14- Registration form
    df = df.ix[0:, 'V16':'V189']     #Slicing useful columns and rows
    df.columns = df.iloc[0]
    df = df.ix[1:,]

    return df

def entriesFormatting(df):
    #df = df.loc[df['V20'] == 'Social Enterprise (SE)']
    #valuesToSlice = ['V16', 'V19', 'V20', 'V21', 'V22', 'V102', 'V105', 'V103', 'V109', 'V112', 'V113', 'V114', 'V115', 'V116', 'V117', 'V119']
    #df = df[valuesToSlice]
    print(df)

    return df

def writeToExcel(df, sheetName):
    writer = ExcelWriter('2015_McGill_Dobson_Cup.xlsx')
    df.to_excel(writer, sheetName, index=False, startrow=5, startcol=3)
    writer.save()



def xlDobson():
    """
    This is a wrapper around fibonacci() to handle all the Excel stuff
    """
    df = pd.read_csv('2015_McGill_Dobson_Cup.csv')

    df = cleanData(df)

    df = entriesFormatting(df)

    #writeToExcel(df, "ProcessedData")


if __name__ == "__main__":
#    if not hasattr(sys, 'frozen'):
#        # The next two lines are here to run the example from Python
#        # Ignore them when called in the frozen/standalone version
#        path = os.path.abspath(os.path.join(os.path.dirname(__file__), '2015_McGill_Dobson_Cup.csv'))
#        Workbook.set_mock_caller(path)
    xlDobson()