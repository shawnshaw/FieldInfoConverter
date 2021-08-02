# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
import xml.etree.ElementTree as ET
import xlsxwriter

sheetPair = {}
barItem = {}

listOfBarItem = []
MenuBarItemPair = {}

ClassNameFieldNoListPair = {}
listFieldNo = []
fieldNo = {}

comboboxInfo = {}
listComboBoxInfo = []
FieldNoComboBoxListPair = {}

resultPair = {}
resultList = []

listWorkSheet = []

sheetName = ''
className = ''

listSheetName = []
listClassName = []

resultDic = {}


def print_hi():
    # Use a breakpoint in the code line below to debug your script.
    mytree = ET.parse('/Users/chenmeirong/PycharmProjects/pythonProject/venv/FieldInfo.xml')
    myroot = mytree.getroot()
    result = xlsxwriter.Workbook('/Users/chenmeirong/PycharmProjects/pythonProject/venv/Result.xlsx')
    worksheet0 = result.add_worksheet('目录')


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
                    worksheetCurrent = result.add_worksheet(sheetName)

                    listWorkSheet.append(worksheetCurrent)

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
                        barItem['menuname'] = j.attrib.get('Name')
                        barItem['id'] = k.attrib.get('ID')
                        if k.attrib.get('ClassName') == '':
                            barItem['className'] = 'N/A'
                        else:
                            barItem['className'] = k.attrib.get('ClassName')
                        barItem['name'] = k.attrib.get('Name')

                        resultList.append([barItem['id'],barItem['className'],barItem['name']])

                        resultDic[j.attrib.get('Name')] = barItem

                        #write to excel
                        row = 0
                        col = 0
                        worksheetCurrent.write(0,0,'id')
                        worksheetCurrent.write(0,1,'ClassName')
                        worksheetCurrent.write(0,2,'name')

                        for list in resultList:
                            #print(list)
                            worksheetCurrent.write(row+1,col,list[col])
                            worksheetCurrent.write(row + 1, col+1, list[col+1])
                            worksheetCurrent.write(row + 1, col+2, list[col+2])
                            row+=1

                        listOfBarItem.append(barItem)
        if root.tag == 'Class':
            for j in root:
                if j.tag == 'ClassName':
                    for k in j:
                        #'className': 'HisMatch', 'name': 'CallOrPutFlag', 'fieldNo': 'CallOrPutFlag', 'chName': '看涨看跌', 'detailInfo': 'CommonCombobox'
                        fieldNo['className'] = j.attrib.get('Name')
                        fieldNo['name'] = k.attrib.get('Name')
                        fieldNo['fieldNo'] = k.attrib.get('FieldName')
                        fieldNo['chName'] = k.attrib.get('ChName')
                        fieldNo['detailInfo'] = k.attrib.get('DetailInfo')
                        #print(fieldNo)
                        listClassName.append(j.attrib.get('Name'))
        if root.tag == 'ComboBoxInfo':
            for j in root:
                if j.tag == 'FieldNo':
                    for k in j:
                        comboboxInfo['fieldNo'] = j.attrib.get('Name')
                        comboboxInfo['enum'] = k.attrib.get('enum')
                        comboboxInfo['description'] = k.attrib.get('description')



                        #print(comboboxInfo)

    print(fieldNo)

    result.close()


     #   print(i.tag)
# Press the green button in the gutter to run the script.
if __name__ == '__main__':
   print_hi()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
