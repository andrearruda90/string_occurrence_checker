import numpy as np 
from pandas import read_excel 
import pandas as pd 
import os


directory = "csvconverter"
parentDir = os.path.expandvars('%systemdrive%/')
path = os.path.join(parentDir,directory)
filesCounter = 0
isPath = os.path.isdir(path)

#creating path if it doesn't exists
if not isPath:
    pass 
    os.mkdir(path)
    print("Directory '% s' created" % directory)

#check if there's any file inside path

for paths in os.listdir(path):
    if os.path.isfile(os.path(path, paths)):
        count +=1

print('File count: ', count)

#df_new = pd.read_csv("C:/Users/andre/Desktop/scdkst00.csv", sep=None) 
#GFG = pd.ExcelWriter("C:/Users/andre/Desktop/scdkst00.xlsx") 
#df_new.to_excel(GFG, index = False) 
  
#GFG.save()

