# The Interface between Python and Excel 
## Warning: Please close the tool before running this code
import os
import numpy as np
import xlwings as xw
import time
import sys
import argparse
from pathlib import Path # Core Python MOdule
#from Tester import Tester

str_inp=[]
str_out=[]
# Please specify the number of inputs you would like to change
num_inp=3
input_udf=np.zeros(num_inp)
input_old=np.zeros(num_inp)
#'Number of Modules'
input_udf[0]=20
#'Hydrogen Production Rate per Module', tonne-H2/hr
input_udf[1]=0.7312
#'Adjusted Operating Capacity Factor' (between 0 and 1)
input_udf[2]=1
print (sys.argv[0])
#file_name=sys.argv[0]
#[xlwings] Write the inputs, run the tool

# # Read the file name of the excel from the tests file
if __name__ == "__main__":
  parser = argparse.ArgumentParser(description="pass the excel file name to python script")
  # Inputs from the user
  parser.add_argument("xlsb_python", help="xlsb_python")
  args = parser.parse_args()
  file_name=args.xlsb_python
  wr=xw.Book(r'./'+ file_name)

sht=wr.sheets['Inputs (Add)']

# store the old value of the parameters
input_old[0]=sht.range('Inp_num_mod').value
input_old[1]=sht.range('Inp_H2_prod_hr').value
input_old[2]=sht.range('Inp_AOCF').value

#reset the input values to zero
sht.range('Inp_num_mod').value=0
sht.range('Inp_H2_prod_hr').value=0
sht.range('Inp_AOCF').value=0

# use the user-defined inputs to check the results
sht.range('Inp_num_mod').value=input_udf[0]
sht.range('Inp_H2_prod_hr').value=input_udf[1]
sht.range('Inp_AOCF').value=input_udf[2]
wr.save()

# Update Step-3 (sensitivity study)
# sht_dash=wr.sheets['Dashboard']
# sht_dash.range(30,14).value=7
# macro1=wr.macro("Module2.runSA_LCOH")
# macro2=wr.macro("Module2.run_SA_two_variable")
# macro3=wr.macro("Module2.run_NPV_LMP")
# macro4=wr.macro("LCOH_SMR_HTSE")
# macro1()
# macro2()
# macro3()
# macro4()
# wr.save()

# [xlwings] read outputs from the tool
#Specify how many outputs are required to compare
output_dashboard=sht.range('Real_H2_prod_d').value
gold_value=input_udf[0]*input_udf[1]*input_udf[2]*24

#Return the old value back to the previous settings
sht.range('Inp_num_mod').value=input_old[0]
sht.range('Inp_H2_prod_hr').value=input_old[1]
sht.range('Inp_AOCF').value=input_old[2]

wr.save()
wr.close()
# wr.app.exit()

#(check if error is less than 0.1%)
if (output_dashboard-gold_value)/gold_value<0.001:
    print ('Test passed for calculating Real hydrogen production rate (tonne/day)')
    sys.exit(0)
else:
    print ('Test failed for calculating Real hydrogen production rate (tonne/day)')
    sys.exit(1)

