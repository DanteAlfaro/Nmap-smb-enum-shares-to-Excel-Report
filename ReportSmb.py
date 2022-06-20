import xlsxwriter
from libnmap.parser import NmapParser, NmapParserException

try:
    parsed = NmapParser.parse_fromfile('scanFPart2.xml')
except NmapParserException as ex:
    parsed = NmapParser.parse_fromfile('scanFPart2.xml', incomplete=True)


workbook = xlsxwriter.Workbook('ReportPar2.xlsx')
worksheet = workbook.add_worksheet()

header_count = 0
headers = ['Ruta', 'Anonymous Access', 'Current User Access']
for header in headers:    
    worksheet.write(0, header_count, header)
    header_count += 1
    
row = 1
col = 0
count = 0   
    
dataProcessed = {}
finalitemList = list()
for host in parsed.hosts:
    try:
        data = host.scripts_results[0]['elements']
    except:    
        pass
    keysData = [*data.keys()]
    
    list_items = list()
    for item in data:
        dataItem = data[item]
        if item != 'account_used' and item != 'note':
            try:
                anonymous_access = dataItem['Anonymous access']
            except:
                anonymous_access = 'N/A'
            try:
                current_user = dataItem['Current user access']
            except:
                current_user = 'N/A'
            list_items = [item, anonymous_access, current_user]
            if list_items not in finalitemList:
                finalitemList.append(list_items)

    
for finalItem in finalitemList:
    local_count = 0
    for item in finalItem:
        worksheet.write(row, local_count, item)
        local_count += 1
    row += 1

   

    
workbook.close()
    


            




        
    