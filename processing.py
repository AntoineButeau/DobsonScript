__author__ = 'abuteau'

import os
import sys
from xlwings import Workbook, Range, Sheet
import pandas as pd




def cleanData():
    data = Range('2015ColoredData.csv', 'A1').table.value
    df = pd.DataFrame(data[2:], columns=data[0])
    numberOfRows = len(df.index)
    numbersOfColumns = len(df.columns)

    df = df[pd.notnull(df['V14'])]
    df = df.ix[:, 'V16':'V189']

def xlDobson():
    """
    This is a wrapper around fibonacci() to handle all the Excel stuff
    """
    # Create a reference to the calling Excel Workbook
    wb = Workbook.caller()

    cleanData()


if __name__ == "__main__":
    if not hasattr(sys, 'frozen'):
        # The next two lines are here to run the example from Python
        # Ignore them when called in the frozen/standalone version
        path = os.path.abspath(os.path.join(os.path.dirname(__file__), '2015_McGill_Dobson_Cup.xlsm'))
        Workbook.set_mock_caller(path)
    xlDobson()