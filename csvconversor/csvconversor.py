import numpy as np 
from pandas import read_excel 
import pandas as pd 
df_new = pd.read_csv("C:/Users/andre/Desktop/scdkst00.csv", sep=None) 
GFG = pd.ExcelWriter("C:/Users/andre/Desktop/scdkst00.xlsx") 
df_new.to_excel(GFG, index = False) 
  
GFG.save()
