#encoding: utf-8 

# add by stefan 20191106 20:14 for develop.
# xls export lua file.

import codecs


def exportLuaFile(filename, list):
    filecontext = "local " + filename + " = {\n"
    filecontext += "values = {"
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

    with codecs.open(filename+'.lua', "w", "utf-8") as target:
        target.write(filecontext)




