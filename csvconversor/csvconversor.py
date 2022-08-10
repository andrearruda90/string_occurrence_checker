from pickle import FALSE
import numpy as np 
from pandas import read_excel 
import pandas as pd 
import os
import time
import glob


directory = "converting/"
parentDir = os.path.expandvars('%systemdrive%/')
path = os.path.join(parentDir,directory) + "converting.csv"
outputfilename = "converted.xlsx"



while(os.path.isfile(path) == False):
    time.sleep(1)

excel_files = glob.glob('/*xlsx*')

for excel_file in excel_files:
    print("Converting '{}'".format(excel_file))
    try:
        df = pd.read_excel(path)
        output = path.split('.')[0]+'.csv'
        df.to_csv(os.path.join(parentDir,directory) + "converted.xlsx" )    
    except KeyError:
        print("  Failed to convert")
#GFdf_new = pd.read_csv(path + filename, sep=None) 
#GFG = pd.ExcelWriter(path + outputfilename)
#GFdf_new.to_excel(GFG, index = False) 

#GFG.save()