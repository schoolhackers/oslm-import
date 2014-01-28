import xlrd
from collections import OrderedDict
import simplejson as json
import glob
thing_list = []
for xls_file in glob.glob("????.xls"):
    #print xls_file
    # Open the workbook and select the first worksheet
    wb = xlrd.open_workbook(xls_file)
    sh = wb.sheet_by_index(0)
    thing = OrderedDict()
    thing['signatur'] = sh.cell(0,2).value
    thing['systematik'] = sh.cell(0,4).value
    thing['autor'] = sh.cell(3,1).value.split(":")[0]
    thing['description'] = ""
    for i in range(4,11):
        try:
            thing['description'] += sh.cell(i,1).value
        except:
            pass
    try:
        thing['isbn'] = sh.cell(11,1).value
    except:
        pass
    thing['type'] = "book"
    thing_list.append(thing)
for xls_file in glob.glob("P-*.xls"):
    #print xls_file
    wb = xlrd.open_workbook(xls_file)
    sh = wb.sheet_by_index(0)
    thing = OrderedDict()
    thing['signatur'] = sh.cell(0,2).value
    thing['description'] = ""
    for i in range(4,11):
        try:
            thing['description'] += sh.cell(i,1).value
        except:
            pass
    thing['type'] = "puzzle"
    thing_list.append(thing)
for xls_file in glob.glob("PC-*.xls"):
    #print xls_file
    # Open the workbook and select the first worksheet
    wb = xlrd.open_workbook(xls_file)
    sh = wb.sheet_by_index(0)
    thing = OrderedDict()
    thing['signatur'] = sh.cell(0,2).value
    thing['description'] = ""
    for i in range(4,7):
        try:
            thing['description'] += sh.cell(i,1).value
        except:
            pass
    thing['type'] = "cdrom"
    thing_list.append(thing)
for xls_file in glob.glob("S-*.xls"):
    #print xls_file
    # Open the workbook and select the first worksheet
    wb = xlrd.open_workbook(xls_file)
    sh = wb.sheet_by_index(0)
    thing = OrderedDict()
    thing['signatur'] = sh.cell(0,2).value
    thing['description'] = ""
    for i in range(3,8):
        try:
            thing['description'] += sh.cell(i,1).value
        except:
            pass
    try:
        thing['isbn'] = sh.cell(11,1).value
    except:
        pass
    thing['type'] = "game"
    thing['age'] = sh.cell(8,2).value
    thing['player'] = sh.cell(9,2).value

    thing_list.append(thing)
j = json.dumps(thing_list)
print j
