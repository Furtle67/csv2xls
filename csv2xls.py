#!/usr/bin/env python3

# coding: utf-8
# Copyright (c) 2016, 2024, Cernus Ltd.  All rights reserved.
#
#   aubrey.blazey@cernus.cloud
#
#   This script combines the csv output of the tenancy  audit tool 
#   into a single XLSX  file with formatting and filters
#   
#   python csv2xls.py <csv_folder> <output.xlsx> [separator]
#
# This script accepts 3 arguments:
#
#   * The folder containg the .csv  files  exported from the tool (/ or \, trailing optional)
#   * The output xlsx filename (.xlsx added if missing, may include a folder)
#   * CSV separator (default ,)
#
# Output: XLSX file with the given name (relative names land in the folder where this tool was run)
#
#   
###############


from pathlib import Path  #cross-platform file access
import pandas as pd  #dataformatting
import sys   #args

MAXWIDTH = 60  #max column width in characters


#CSV folder and output filename passed in
if len(sys.argv) > 2:
    csvLoc=Path(sys.argv[1]).expanduser()
    xlsxName=Path(sys.argv[2]).expanduser()
    if xlsxName.suffix.lower() != ".xlsx":
        xlsxName=xlsxName.with_name(xlsxName.name+".xlsx")
    seperator=sys.argv[3] if len(sys.argv) > 3 else ","   #set seperator to comma if not passed in
    #shells pass \t as backslash + t, not a tab - convert so tab separated files work
    if seperator in ("\\t", "tab"):
        seperator="\t"
    print(f"CSV location : {csvLoc}")
    print(f"Output file : {xlsxName}")
    print(f"sep : {seperator!r}")  #!r shows tab as '\t'

    #fail early - don't write an empty xlsx
    if not csvLoc.is_dir():
        print(f"ERROR : CSV folder {csvLoc} not found")
        sys.exit(1)
    csvFiles = sorted(csvLoc.glob('*.csv'))
    if not csvFiles:
        print(f"ERROR : no .csv files found in {csvLoc}")
        sys.exit(1)

    #workbook created on first csv with content - ExcelWriter empties the file as soon as it opens it
    writer = None
    totalDupes = 0  #duplicate rows removed across all csvs

    for csvfile in csvFiles:
        print(csvfile)  #debug
        #max sheet name 31 chars
        #oci
        sheetName = csvfile.stem  #filename only, no folder or .csv - works with / or \
        sheetName = sheetName.replace("identity","id")
        sheetName = sheetName.replace("network","net")
        sheetName = sheetName.replace("object_storage","os")
        sheetName = sheetName.replace("security","sec")
        sheetName = sheetName.replace("block","blk")
        sheetName = sheetName.replace("database_goldengate","goldengate")
        #excel max sheet name 31 chars
        sheetName = sheetName[:31]

        print(sheetName)

        #csv into dataframe
        #dtype=str keeps values as-is, eg leading zeros in AWS account IDs, long IDs not rounded
        try:
            df_new = pd.read_csv(csvfile,sep=seperator,dtype=str)
        except pd.errors.EmptyDataError:
            print(f"WARNING : {csvfile} is empty - skipped")
            continue
        #clean
        rowsBefore = len(df_new)
        df_new.drop_duplicates(inplace=True)
        dupes = rowsBefore - len(df_new)
        if dupes:
            print(f"Duplicate rows removed : {dupes}")
            totalDupes += dupes
        #only drop empty columns when there is data - header-only csv would lose all columns
        if len(df_new):
            df_new.dropna(axis=1,how="all",inplace=True)
        
        #create workbook
        # choose writer openpyxl | xlsxwriter
        #writer = pd.ExcelWriter(tenancy+".xlsx", engine="openpyxl",mode='w')
        if writer is None:
            writer = pd.ExcelWriter(xlsxName, engine="xlsxwriter",mode='w')

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
        #filter across all columns and rows of the data (row, col zero based)
        if len(df_new.columns):
            worksheet.autofilter(0, 0, len(df_new), len(df_new.columns) - 1)
        worksheet.freeze_panes(1, 0)  #keep header row on screen when scrolling

        # Size columns to fit data. Header gets +4 so the filter arrow doesn't cover it,
        # capped at MAXWIDTH so long text columns stay readable
        for col_num, col in enumerate(df_new.columns):
            dataLen = df_new[col].astype(str).str.len().max() if len(df_new) else 0
            width = min(max(len(str(col)) + 4, dataLen + 2), MAXWIDTH)
            worksheet.set_column(col_num, col_num, width)


    #all csvs skipped as empty - workbook never created, existing xlsx untouched
    if writer is None:
        print(f"ERROR : all .csv files in {csvLoc} were empty - no xlsx written")
        sys.exit(1)

    writer.close()  #close  (save is deprecated)
    print(f"Total duplicate rows removed : {totalDupes}")
    print(f"Combined Cloud Infrastructure Document  : {xlsxName}")    

else:
    print("Please pass in the folder containing CSV file inputs,  the filename for the xlsx output , and optionally the file seperator (default ,)")




