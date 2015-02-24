__author__ = 'abuteau'

import os
import sys
from xlwings import Workbook, Range, Sheet
import pandas as pd
from pandas import ExcelWriter



def cleanData(df):

    df = df[pd.notnull(df['V14'])]   # Removing empty entries from V14- Registration form
    df = df.ix[0:, 'V16':'V189']     #Slicing useful columns and rows

    return df

def entriesFormatting(df):
    valuesToSlice = ['V16', 'V19', 'V20', 'V21', 'V22', 'V102', 'V105', 'V103', 'V109', 'V112', 'V113', 'V114', 'V115', 'V116', 'V117', 'V119', 'V120', 'V123', 'V121', 'V127', 'V130', 'V131', 'V132', 'V133', 'V134', 'V135', 'V137', 'V138', 'V141', 'V139', 'V145', 'V148', 'V149', 'V150', 'V151', 'V152', 'V153', 'V155', 'V156', 'V158', 'V157', 'V162', 'V165', 'V166', 'V167', 'V168', 'V169', 'V170', 'V172', 'V173', 'V175', 'V174', 'V179', 'V182', 'V183', 'V184', 'V185', 'V186', 'V187', 'V189']
    df = df[valuesToSlice]

    return df

def writeEntriesToExcel(df):

    dfSE = df.loc[df['V20'] == 'Social Enterprise (SE)']
    dfIDE = df.loc[df['V20'] == 'Innovation Driven Enterprise (IDE)']
    dfSME = df.loc[df['V20'] == 'Small & Medium Enterprise (SME)']

    with ExcelWriter('2015_McGill_Dobson_Cup.xlsx') as writer:
        dfSE.to_excel(writer, 'SE', index=False, startrow=5, startcol=3)
        dfIDE.to_excel(writer, 'IDE', index=False, startrow=5, startcol=3)
        dfSME.to_excel(writer, 'SME', index=False, startrow=5, startcol=3)

def xlDobson():
    df = pd.read_csv('2015_McGill_Dobson_Cup.csv')

    df = cleanData(df)

    df = entriesFormatting(df)

    writeEntriesToExcel(df)


if __name__ == "__main__":
    xlDobson()