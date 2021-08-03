# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
import xml.etree.ElementTree as ET
import xlsxwriter

MenuBarItemPair = {}


resultPair = {}
resultList = []

listWorkSheet = []

sheetName = ''
className = ''
sheetPair = {}
listSheetName = []
listMenuName = []
listClassName = []
listFieldNo = []
listDetailInfo = []

result = {}
resultExcel = xlsxwriter.Workbook('C:/Users/Ruanyueying/PycharmProjects/Result.xlsx')
worksheet0 = resultExcel.add_worksheet('目录')

def print_hi():
    # Use a breakpoint in the code line below to debug your script.
    mytree = ET.parse('C:/Users/Ruanyueying/PycharmProjects/FieldInfo.xml')
    myroot = mytree.getroot()

    for root in myroot:
        print(root.tag)
        if root.tag == 'NavBar':
            for j in root:
                if j.tag == 'BarControl':
                    # Bar menu {'ID': '110000', name : '基础数据'
                    sheetPair['id'] = j.attrib.get('ID')
                    sheetPair['name'] = j.attrib.get('Name')

                    sheetName = j.attrib.get('Name')

                    listSheetName.append(sheetName)

                    # New Current sheet and write in the latter block later

                    #write to Excel
                    row = 0
                    col = 0

                    resultPair[sheetPair['id']] = sheetPair['name']

                    for key,value in resultPair.items():
                        worksheet0.write(0,0,'id')
                        worksheet0.write(0,1,'name')
                        worksheet0.write(row+1,col,key)
                        worksheet0.write(row+1,col+1,value)
                        #worksheet0.write(row+1,1,sheetPair[key])
                        row += 1

                    # BarItem {'menuname': '客户管理', 'id': '120030', 'className': 'AccountPartyInfo', 'name': '客户关系人'}
                    for k in j:
                        barItem = {}
                        barItem['menuname'] = j.attrib.get('Name')
                        barItem['id'] = k.attrib.get('ID')
                        barItem['name'] = k.attrib.get('Name')
                        if k.attrib.get('ClassName') == '':
                            if barItem['name'] == '客户信息':
                                barItem['className'] = 'AccountInfo'
                            else:
                                barItem['className'] = 'N/A'
                        else:
                            barItem['className'] = k.attrib.get('ClassName')

                        print(barItem)
                        listClassName.append(barItem)


                        # #write to excel
                        # row = 0
                        # col = 0
                        # worksheetCurrent.write(0,0,'id')
                        # worksheetCurrent.write(0,1,'ClassName')
                        # worksheetCurrent.write(0,2,'name')
                        #
                        # for list in resultList:
                        #     #print(list)
                        #     worksheetCurrent.write(row+1,col,list[col])
                        #     worksheetCurrent.write(row + 1, col+1, list[col+1])
                        #     worksheetCurrent.write(row + 1, col+2, list[col+2])
                        #     row+=1

        if root.tag == 'Class':
            for j in root:
                if j.tag == 'ClassName':
                    for k in j:
                        #'className': 'HisMatch', 'name': 'CallOrPutFlag', 'fieldNo': 'CallOrPutFlag', 'chName': '看涨看跌', 'detailInfo': 'CommonCombobox'
                        fieldNo = {}
                        fieldNo['className'] = j.attrib.get('Name')
                        fieldNo['name'] = k.attrib.get('Name')
                        fieldNo['fieldNo'] = k.attrib.get('FieldName')
                        fieldNo['chName'] = k.attrib.get('ChName')
                        fieldNo['detailInfo'] = k.attrib.get('DetailInfo')
                        listFieldNo.append(fieldNo)

        if root.tag == 'ComboBoxInfo':
            for j in root:
                if j.tag == 'FieldNo':
                    for k in j:
                        comboboxInfo = {}
                        comboboxInfo['fieldNo'] = j.attrib.get('Name')
                        comboboxInfo['enum'] = k.attrib.get('enum')
                        comboboxInfo['description'] = k.attrib.get('description')
                        listDetailInfo.append(comboboxInfo)
   # result.close()

def generateResult():
#ClassName: {'menuname': '基础数据', 'id': '110010', 'className': 'CurrencyGroupInfo', 'name': '币种组信息'}
# FieldNo: {'className': 'UserTrustDevice', 'name': 'OperateTime', 'fieldNo': 'OperateTime', 'chName': '操作时间', 'detailInfo': 'OperateTime'}
#comboboxInfo：{'fieldNo': 'ImportFileDesc', 'enum': 'TD', 'description': '1. TD文件，文件名中包含关键字TD，扩展名为.txt，BIG5编码'}
    for sheetName in listSheetName:
        listTemp = []
        for className in listClassName:
            if className['menuname'] == sheetName:
                for fieldNo in listFieldNo:
                    if fieldNo['className'] == className['className']: #and fieldNo['detailInfo'] == 'CommonCombobox' :
                        for comboboxInfo in listDetailInfo:
                            if fieldNo['fieldNo'] == comboboxInfo['fieldNo']:
                                dicTemp = {}
                                dicTemp['id'] = className['id']
                                dicTemp['className'] = className['className']
                                dicTemp['name'] = className['name']
                                dicTemp['fieldNo'] = fieldNo['fieldNo']
                                dicTemp['chName'] = fieldNo['chName']
                                dicTemp['enum'] = comboboxInfo['enum']
                                dicTemp['description'] = comboboxInfo['description']
                                listTemp.append(dicTemp)

            result[sheetName] = listTemp

def writeResultToExcel():
    #'id': '110040', 'className': 'CommodityInfo', 'name': '品种信息', 'fieldNo': 'OptionType', 'chName': '期权类型',
    # 'enum': 'CNY', 'description': '在岸人民币'
    for key in result.keys():
        workSheetMenu = resultExcel.add_worksheet(key)
        # Header
        workSheetMenu.write(0,0,'id')
        workSheetMenu.write(0,1,'className')
        workSheetMenu.write(0,2,'name')
        workSheetMenu.write(0,3,'fieldNo')
        workSheetMenu.write(0,4,'chName')
        workSheetMenu.write(0,5,'enum')
        workSheetMenu.write(0,6,'description')
        row = 1
        for list in result[key]:
            workSheetMenu.write(row, 0, list['id'])
            workSheetMenu.write(row, 1, list['className'])
            workSheetMenu.write(row, 2, list['name'])
            workSheetMenu.write(row, 3, list['fieldNo'])
            workSheetMenu.write(row, 4, list['chName'])
            workSheetMenu.write(row, 5, list['enum'])
            workSheetMenu.write(row, 6, list['description'])
            row += 1



if __name__ == '__main__':
   print_hi()
   generateResult()
   writeResultToExcel()
   resultExcel.close()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
