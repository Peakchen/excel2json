#encoding: utf-8 

# add by stefan 20191106 20:14 for develop.
# xls export json file.

import json
import codecs

def exportJsonFile(filename, list):
    exportf = json.dumps(list, ensure_ascii=False) 
    with codecs.open(filename+'.json', "w", "utf-8") as target:
        target.write(exportf)



