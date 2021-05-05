import pandas as pd 
import numpy as np 
import os
import sys
import glob

#--------------------------------------------------------------------------------------------
#	Author: Nikhil M
#	Created On: 05/03/2021
#	Use: python ConvertCSVToExcel.py 'argument1' argument2 > Logfile along with path
#	Arguments: 1. CSV_FILE wildcard
#		   2. CSV FILE PATH; TARGET WILL BE CREATED HERE
#--------------------------------------------------------------------------------------------


if len(sys.argv)!=2:
	print('This script requires 2 arguments. One for the filename and another for path')

CSV_FILE=sys.argv[1]  
PATH=sys.argv[2]

print('First argument is CSV_FILE:'+CSV_FILE)
print('Second argument is PATH:'+PATH)

os.chdir(PATH)
print('Changed directory')

WILDCARD_PATH=PATH+'/'+CSV_FILE
print('Wild card path:'+WILDCARD_PATH)

for fname in glob.glob(WILDCARD_PATH):
	FILENAME_SIZE=len(fname)
	TGT_FILE=fname[:FILENAME_SIZE-4] + '.xlsx'
	print('Target Excel Filename will be:'+TGT_FILE)

	df = pd.read_csv(fname,encoding='iso-8859-1',error_bad_lines=False,index_col=False,header=None)

	GFG = pd.ExcelWriter(TGT_FILE) 
	df.to_excel(GFG, index = False,header=None) 
	  
	GFG.save()
	print('Target excel file '+TGT_FILE+' is on the server')
print('For loop ended')

