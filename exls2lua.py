#encoding: utf-8 

# add by stefan 20191106 20:14 for develop.
# xls export lua file.

import xlrd 
from collections import OrderedDict
import json
import codecs
import glob
import os,sys
import os.path
import re

reload(sys)
sys.setdefaultencoding("utf8")

totaltabName = "总揽"
totaltabs = {}

def loopOpenXld2Lua():
    print os.getcwd()
    for file in glob.glob("*.xls"):
        wb = xlrd.open_workbook(file)
        sheets = wb.sheets()
        #print "show sheets list: "
        for sheet in sheets:
            exportSheet(sheet)

def exportSheet(sheet):
    #print sheet.name
    shn = sheet.name
    shn = shn.encode("UTF-8")  
    if shn == totaltabName:
        exportTotalTab(sheet)
    else:
        print shn
        exportLua(sheet)

#
# 获取总揽数据
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

# 开始导出xls数据到json文件中
def exportLua(sheet):
    convertlua_list = []
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

        if rowidx == 2 and row[1].value != shn: # 获取表名，是否可导出
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
                    if len(rowdata) == totallimit and rowdata[strcolidx] == 's' :
                        continue
            
            if colitemtypes.has_key('4') == True and rowidx >= 7:
                rowdata = colitemtypes['4']
                if rowdata.has_key(strcolidx) == True:
                    if len(rowdata) == totallimit:
                        if rowdata[strcolidx] == 'Integer':
                            colval = int(colval)
                        elif rowdata[strcolidx] == 'String':
                            print colval
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
            convertlua_list.append(colitemtypes[strRowidx])

    createRow = colitemtypes['2'] # 获取表名，是否可导出
    if totaltabs.has_key(createRow['2']) == True and len(convertlua_list) > 0:
        print "find it, it can be export."
        #开始导出lua数据
        beginExportLua(shn, convertlua_list)


def beginExportLua(filename, list):
    print list
    filecontext = "local " + filename + " = {\n"
    filecontext += "values = {"
    #print filecontext
    rowidx = 0
    for items in list:
        rowidx = rowidx + 1
        colidx = 0
        filecontext += "{"
        for item in items:
            colidx = colidx + 1
            value = items[item]
            if isinstance(value, int):
                strvalue = '{}'.format(value)
                filecontext += item + "=" + strvalue
            elif isinstance(value, str):
                print value
                value = value.encode("UTF-8") 
                filecontext += item + "=" + '"' +value + '"'

            if colidx == len(items) and rowidx < len(list):
                filecontext += "},\n"
            elif colidx == len(items) and rowidx == len(list):
                filecontext += "}\n"
            else:
                filecontext += ","

    filecontext += "}\n"
    filecontext += "}\n"

    print filecontext
    with codecs.open(filename+'.lua', "w", "utf-8") as target:
        target.write(filecontext)

loopOpenXld2Lua()



