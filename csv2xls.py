#!/bin/python3

# coding: utf-8
# Copyright (c) 2016, 2024, Cernus Ltd.  All rights reserved.
#
#   aubrey.blazey@cernus.cloud
#
#   This script combines the csv output of the tenancy  audit tool 
#   into a single XLSX  file with formatting and filters
#   
#
#
# This script accepts 2 arguments:
#
#   * The folder containg the .csv  files  exported from the tool (without trailing /)
#   * The tenancy name (or xls target filename)
#
# Output: XLS  file with  Tenancy and Date in the folder where this tool was run
#
#    AWS version
#   
###############


import os #file access
import glob  #file access
import pandas as pd  #dataformatting
import sys   #args
from datetime import date


############################################
# set DATE  variable for file naming
############################################
vToday = date.today()
VDATE = vToday.strftime("%Y%m%d")

#Tenancy name passed in
if len(sys.argv) > 2:
    csvLoc=sys.argv[1]
    tenancy=sys.argv[2]
    seperator=sys.argv[3]
    print(f"CSV location : {csvLoc}")
    print(f"Tenancy : {tenancy}")
    print(f"sep : {seperator}")
    fileRem=csvLoc
    print(f"File Path  : {fileRem}  will be removed from the glob")
    # cloud infra document ?  Cloud Infra composition  | status? ? inventory? 
    xlsxName=tenancy+"_CID_"+VDATE+".xlsx"
    #create workbook
    # choose writer openpyxl | xlsxwriter
    #writer = pd.ExcelWriter(tenancy+".xlsx", engine="openpyxl",mode='w')    
    writer = pd.ExcelWriter(xlsxName, engine="xlsxwriter",mode='w')    

    for csvfile in sorted(glob.glob(os.path.join(csvLoc, '*.csv'))):
        print(csvfile)  #debug
        #max sheet name 31 chars
        #oci
        sheetName = csvfile.replace(fileRem,"")
        sheetName = sheetName.replace(".csv", "")
        sheetName = sheetName.replace("identity","id")
        sheetName = sheetName.replace("network","net")
        sheetName = sheetName.replace("object_storage","os")
        sheetName = sheetName.replace("security","sec")
        sheetName = sheetName.replace("identity","id")
        sheetName = sheetName.replace("block","blk")
        sheetName = sheetName.replace("database_goldengate","goldengate")
        # AWS
        sheetName = sheetName.replace("block","blk")
        
        
        print(sheetName)

        #csv into dataframe
        df_new = pd.read_csv(csvfile,sep=seperator)
        #clean
        df_new.drop_duplicates(inplace=True)
        df_new.dropna(axis=1,how="all",inplace=True)
        
        #create xls
        df_new.to_excel(writer, sheet_name=str(sheetName),header=True,index=False)
        #now access the worksheet and add formatting
        # Get the xlsxwriter workbook and worksheet objects.
        # Autofit the worksheet.
        workbook = writer.book
        worksheet = writer.sheets[sheetName]

        # Add a header format.
        header_format = workbook.add_format({
            "bold": True,
            "align": "left",
            "valign": "top",
            "font_color": "white",
            "bg_color": "#A9A9A9",
            "border": 1})   
        

        # Write the column headers with the defined format. write(row, col, *args)
        for col_num, value in enumerate(df_new.columns.values):
            worksheet.write(0, col_num , value, header_format)
        worksheet.autofilter("A1:Z1")
        worksheet.autofit()


    writer.close()  #close  (save is deprecated)
    print(f"Combined Cloud Infrastructure Document  : {xlsxName}")    

else:
    print("Please pass in CSV  file location and Tenancy Name")




