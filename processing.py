__author__ = 'abuteau'

import os
import sys
from xlwings import Workbook, Range, Sheet
import pandas as pd
from pandas import ExcelWriter
import numpy as np


def cleanData(df):

    df = df[pd.notnull(df['V14'])]   # Removing empty entries from V14- Registration form
    df = df.ix[:, 'V16':'V189']     #Slicing useful columns and rows

    return df

def concat(*args):
    strs = [str(arg) for arg in args if not pd.isnull(arg)]
    return ','.join(strs) if strs else np.nan

def entriesFormatting(df):
    np_concat = np.vectorize(concat)

    dfStartupBasicInfo = df.ix[:,'V16':'V22']
    dfMember1 = df.ix[:,'V102':'V119']
    dfMember2 = df.ix[:,'V120':'V137']
    dfMember3 = df.ix[:, 'V138':'V155']
    dfMember4 = df.ix[:, 'V156':'V172']
    dfMember5 = df.ix[:,'V173':'V189']
    dfStartupMoreInfo = df.ix[:,'V46':'V101']
    dfHowWeMet = df.ix[:,'V23':'V45']

    dfHowWeMet['V23'] = np_concat(dfHowWeMet['V23'], dfHowWeMet['V24'], dfHowWeMet['V25'], dfHowWeMet['V26'], dfHowWeMet['V27'], dfHowWeMet['V28'], dfHowWeMet['V29'], dfHowWeMet['V30'], dfHowWeMet['V31'], dfHowWeMet['V32'], dfHowWeMet['V33'], dfHowWeMet['V34'], dfHowWeMet['V35'], dfHowWeMet['V36'], dfHowWeMet['V37'], dfHowWeMet['V38'], dfHowWeMet['V39'], dfHowWeMet['V40'], dfHowWeMet['V41'], dfHowWeMet['V42'], dfHowWeMet['V43'], dfHowWeMet['V44'], dfHowWeMet['V45'])

    df = dfStartupBasicInfo.join(dfHowWeMet['V23']).join(dfMember1).join(dfMember2).join(dfStartupMoreInfo)

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