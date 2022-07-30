from pickle import FALSE
import numpy as np 
from pandas import read_excel 
import pandas as pd 
import os
import time


directory = "converting/"
parentDir = os.path.expandvars('%systemdrive%/')
path = os.path.join(parentDir,directory)
filename = "converting.csv"
outputfilename = "converted.xlsx"



while(os.path.isfile(path + filename) == False):
    time.sleep(1)

GFdf_new = pd.read_csv(path + filename, sep=None) 
GFG = pd.ExcelWriter(path + outputfilename)
GFdf_new.to_excel(GFG, index = False) 

GFG.save()