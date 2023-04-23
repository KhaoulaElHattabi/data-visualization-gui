import pandas as pd
import numpy as np
cols = [1, 2, 3]
workbook = pd.read_excel('C:\\Users\\Khaaoula\\Desktop\\Stelia\\PFA_Stelia\\input\\SUIVI_POLYME.xlsx',sheet_name='AMN',usecols=cols)
workbook=workbook.dropna()
print(workbook.head())