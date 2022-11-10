import pandas as pd
import numpy as np

def replacecolumvalues(values_to_replace,dataframe):
    for col in dataframe.columns:
        dataframe[col].replace(to_replace=values_to_replace, regex=True,inplace=True)
    return dataframe





