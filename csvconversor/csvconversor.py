import numpy as np 
from pandas import read_excel 
import pandas as pd
import os 
dir_path = os.path.dirname(str(os.path.realpath("csvconversor.py") + "scdkst00.xlsx"))
df_new = pd.read_csv("C:/Users/andre.arruda/Desktop/scdkst00.csv", sep=None) 
GFG = pd.ExcelWriter(dir_path) 
df_new.to_excel(GFG, index = False) 
 
GFG.save()
