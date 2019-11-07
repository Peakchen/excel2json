#encoding: utf-8 

# add by stefan 20191107 10:24 for develop.
# xls export json file.

import xlrd 
from collections import OrderedDict
import json
import codecs
import glob
import os,sys
import os.path
import re

import sys
reload(sys)
sys.setdefaultencoding("utf8")

from define import ExportFiles, ExportFilter
from exls2json import exportJsonFile
from exls2lua import exportLuaFile

totaltabName = "总揽"
totaltabs = {}

# purpose: read file and start export
# param1: need export aim file.
# param2: filter condition.
def loopExportXld(dstfile, filter):
    print os.getcwd()
    for file in glob.glob("*.xls"):
        wb = xlrd.open_workbook(file)
        sheets = wb.sheets()
        #print "show sheets list: "
        for sheet in sheets:
            exportSheet(sheet, dstfile, filter)

def exportSheet(sheet, dstfile, filter):
    #print sheet.name
    shn = sheet.name
    shn = shn.encode("UTF-8")  
    if shn == totaltabName:
        exportTotalTab(sheet)
    else:
        print shn
        export2File(sheet, dstfile, filter)

# purpose: export total table from others exported.
# param: tab sheet.

def exportTotalTab(sheet):
    print sheet.name
    rows = sheet.get_rows()
    rowidx = 0
    for row in rows:
        rowidx = rowidx + 1
        if rowidx == 1:
            continue
        
        tablename = row[0].value
        if tablename == '':
            continue

        isexport = row[1].value
        begincol = int(row[4].value)
        endcol = int(row[5].value)
        totaltabs[tablename] = {'isexport':isexport, 'begincol': begincol, 'endcol': endcol} 
        #print totaltabs

# begin export operation.
# param1: tab sheet.
# param2: need export aim file.
# param3: filter condition.
def export2File(sheet, dstfile, filter):
    convertlist = []
    shn = sheet.name
    shn = shn.encode("UTF-8")  
    if totaltabs.has_key(shn) == False:
        return

    expitem = totaltabs[shn]
    #print expitem
    rows = sheet.get_rows()
    rowidx = 0
    colitemtypes = OrderedDict()
    totallimit = expitem['endcol'] - 1
    for row in rows:
        rowidx = rowidx + 1
        if row[0].value == '':
            continue

        if row[0].value == 'CLOSE':
            break

        if rowidx == 2 and row[1].value != shn: # if not exist, then exist.
            return

        #print row
        colidx = 0
        for col in row:
            colval = col.value
            if colval == '' and colidx == 1:
                continue
        
            colidx = colidx + 1
            strRowidx = str(rowidx)
            strcolidx = str(colidx)

            if totallimit < colidx:
                continue

            if colitemtypes.has_key('3') == True and rowidx >= 7:
                rowdata = colitemtypes['3']
                if rowdata.has_key(strcolidx) == True:
                    totallimit = expitem['endcol'] - 1
                    if len(rowdata) == totallimit and rowdata[strcolidx] == filter:
                        continue
            
            if colitemtypes.has_key('4') == True and rowidx >= 7:
                rowdata = colitemtypes['4']
                if rowdata.has_key(strcolidx) == True:
                    if len(rowdata) == totallimit:
                        if rowdata[strcolidx] == 'Integer':
                            colval = int(colval)
                        elif rowdata[strcolidx] == 'String':
                            colval = str(colval)
                        elif rowdata[strcolidx] == 'Float':
                            colval = float(colval)   

            if isinstance(colval, str):
                colval = colval.encode("UTF-8") 

            if colitemtypes.has_key(strRowidx) == False:
                if rowidx >= 7 and colidx > 1:
                    colitem = colitemtypes['6']
                    if totaltabs.has_key(strRowidx) == False:
                        colitemtypes[strRowidx] = {colitem[strcolidx]: colval}
                    else:
                        colitemtypes[strRowidx].update({colitem[strcolidx]: colval})  
                elif rowidx < 7:
                    colitemtypes[strRowidx] = {strcolidx: colval}
            else:
                if rowidx >= 7 and colidx > 1:
                    colitem = colitemtypes['6']
                    colitemtypes[strRowidx].update({colitem[strcolidx]: colval})  
                elif rowidx < 7:
                    colitemtypes[strRowidx].update({strcolidx: colval})  

        if rowidx >= 7 and colitemtypes.has_key(strRowidx):
            convertlist.append(colitemtypes[strRowidx])

    createRow = colitemtypes['2'] # get table file, judge it can be export.
    if totaltabs.has_key(createRow['2']) == True and len(convertlist) > 0:
        print "find it, it can be export."
        if dstfile == "Json":
            exportJsonFile(shn, convertlist)
        elif dstfile == "Lua":
            exportLuaFile(shn, convertlist)


def _main_():
    for file in ExportFiles:
        filter = ExportFilter[file]
        loopExportXld(file, filter)


_main_()
