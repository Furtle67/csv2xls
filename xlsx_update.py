#!/usr/bin/env python3

# coding: utf-8
# Copyright (c) 2016, 2024, Cernus Ltd.  All rights reserved.
#
#   aubrey.blazey@cernus.cloud
#
#   This script copies the sheets from a fresh xlsx (output of csv2xls.py)
#   into an existing target xlsx, replacing sheets with the same name
#
#   python xlsx_update.py <fresh.xlsx> <target.xlsx>
#
# This script accepts 2 arguments:
#
#   * The fresh xlsx file (from csv2xls.py)
#   * The target xlsx file to update
#
# Behaviour:
#   * target is backed up first to <target>_backup_YYYYMMDD_HHMMSS.xlsx
#   * sheets with a matching name (case-insensitive, as Excel) are replaced in the same position
#   * sheets not in the target are added at the end
#   * PROTECTED sheets (eg Team) are never replaced
#   * other target sheets are left untouched - formulas (eg XLOOKUP) are kept
#
###############


from pathlib import Path  #cross-platform file access
from datetime import datetime  #backup filename
import shutil  #backup copy
import sys   #args
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MAXWIDTH = 60  #max column width in characters - same as csv2xls.py
PROTECTED = {"team"}  #sheet names never replaced (lowercase)

#header format - same as csv2xls.py
thin = Side(style="thin")
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="A9A9A9")
HEADER_ALIGN = Alignment(horizontal="left", vertical="top")
HEADER_BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


#fresh and target filenames passed in
if len(sys.argv) > 2:
    freshName=Path(sys.argv[1]).expanduser()
    targetName=Path(sys.argv[2]).expanduser()
    print(f"Fresh file : {freshName}")
    print(f"Target file : {targetName}")

    #fail early before touching the target
    for f in (freshName, targetName):
        if not f.is_file():
            print(f"ERROR : {f} not found")
            sys.exit(1)
    if freshName.resolve() == targetName.resolve():
        print("ERROR : fresh and target are the same file")
        sys.exit(1)

    freshWb = load_workbook(freshName, read_only=True)
    targetWb = load_workbook(targetName)  #not read_only - keeps formulas as written

    #backup target before any change
    backupName = targetName.with_name(f"{targetName.stem}_backup_{datetime.now():%Y%m%d_%H%M%S}{targetName.suffix}")
    shutil.copy2(targetName, backupName)
    print(f"Backup : {backupName}")

    for srcWs in freshWb.worksheets:
        sheetName = srcWs.title

        if sheetName.lower() in PROTECTED:
            print(f"WARNING : {sheetName} is protected - skipped")
            continue

        #excel sheet names are case-insensitive - match the same way
        existing = {name.lower(): name for name in targetWb.sheetnames}
        if sheetName.lower() in existing:
            oldWs = targetWb[existing[sheetName.lower()]]
            pos = targetWb.index(oldWs)
            targetWb.remove(oldWs)
            ws = targetWb.create_sheet(sheetName, pos)
            print(f"Replaced : {sheetName}")
        else:
            ws = targetWb.create_sheet(sheetName)
            print(f"Added : {sheetName}")

        #copy values as-is - text stays text, leading zeros kept
        for row in srcWs.iter_rows(values_only=True):
            ws.append(row)

        if ws.max_row < 1 or ws.max_column < 1 or ws.cell(1, 1).value is None:
            continue  #nothing to format

        #header format
        for cell in ws[1]:
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = HEADER_ALIGN
            cell.border = HEADER_BORDER

        ws.auto_filter.ref = ws.dimensions  #filter across all columns and rows of the data
        ws.freeze_panes = "A2"  #keep header row on screen when scrolling

        # Size columns to fit data. Header gets +4 so the filter arrow doesn't cover it,
        # capped at MAXWIDTH so long text columns stay readable
        for col_num, col in enumerate(ws.iter_cols(values_only=True), start=1):
            header = str(col[0]) if col[0] is not None else ""
            dataLen = max((len(str(v)) for v in col[1:] if v is not None), default=0)
            width = min(max(len(header) + 4, dataLen + 2), MAXWIDTH)
            ws.column_dimensions[get_column_letter(col_num)].width = width

    freshWb.close()

    try:
        targetWb.save(targetName)
    except PermissionError:
        print(f"ERROR : cannot write {targetName} - is it open in Excel? Target unchanged, backup at {backupName}")
        sys.exit(1)
    print(f"Updated : {targetName}")

else:
    print("Please pass in the fresh xlsx (from csv2xls.py) and the target xlsx to update")
