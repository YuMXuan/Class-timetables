# coding: utf-8
#!/usr/bin/python

import sys
import json
import xlrd
import importlib

def main():
    # 读取 excel 文件
    workbook = xlrd.open_workbook('y1.xlsx')

    headStr = '{\n"classInfo":[\n'
    tailStr = ']\n}'
    itemHeadStr = '{\n'
    itemTailStr = '\n}'
    classInfoStr1 = ''
    classInfoStr1 += headStr
    classInfoStr2 = ''
    classInfoStr2 += headStr

    week = 0

    for s in range(0,len(workbook.sheets())):
        week = week+1
        print(week*5)
        sheet = workbook.sheet_by_index(s)
        numOfRow = sheet.nrows
        numOfCol = sheet.ncols
        weekday = 0
        for ci in range(3,13,2):
            weekday = weekday+1
            sessionlist1 = []
            sessionlist2 = []
            module1 = []
            module2 = []
            for ri in range(3,19):
                if (sheet.cell(ri, ci).value is None or sheet.cell(ri, ci).value == ''):
                    if get_merged_cells_value(sheet, ri, ci):
                        module1.append(str(get_merged_cells_value(sheet, ri, ci)))
                        sessionlist1.append(str(int(ri-2)))
                else:
                    module1.append(str(sheet.cell(ri, ci).value))
                    sessionlist1.append(str(int(ri-2)))
            cii = ci+1
            for ri in range(3,19):
                if (sheet.cell(ri, cii).value is None or sheet.cell(ri, cii).value == ''):
                    if get_merged_cells_value(sheet, ri, cii):
                        module2.append(str(get_merged_cells_value(sheet, ri, cii)))
                        sessionlist2.append(str(int(ri-2)))
                else:
                    module2.append(str(sheet.cell(ri, cii).value))
                    sessionlist2.append(str(int(ri-2)))
            module1 = map(lambda x: x.replace('1-7', 'G1-7').replace('\n', '').replace('（', '(').replace('，', ','), module1)
            module2 = map(lambda x: x.replace('8-14', 'G8-14').replace('\n', '').replace('）', ')').replace('，', ','), module2)
            i = 0
            for className1 in module1:
                itemClassInfoStr1 = ""
                itemClassInfoStr1 += itemHeadStr + '"className":"' + className1 + '",\n'
                itemClassInfoStr1 += '"week":' + str(week) + ',\n'
                itemClassInfoStr1 += '"weekday":' + str(weekday) + ',\n'
                itemClassInfoStr1 += '"session":' + sessionlist1[i]
                itemClassInfoStr1 += itemTailStr
                itemClassInfoStr1 += ","
                classInfoStr1 += itemClassInfoStr1
                i += 1
            i = 0
            for className2 in module2:
                itemClassInfoStr2 = ""
                itemClassInfoStr2 += itemHeadStr + '"className":"' + className2 + '",\n'
                itemClassInfoStr2 += '"week":' + str(week) + ',\n'
                itemClassInfoStr2 += '"weekday":' + str(weekday) + ',\n'
                itemClassInfoStr2 += '"session":' + sessionlist2[i] 
                itemClassInfoStr2 += itemTailStr
                itemClassInfoStr2 += ","
                classInfoStr2 += itemClassInfoStr2
                i += 1
    classInfoStr1 = classInfoStr1.strip(",")
    classInfoStr1 += tailStr
    classInfoStr2 = classInfoStr2.strip(",")
    classInfoStr2 += tailStr
    with open('conf_classInfo_G1-7.json','w') as f:
	    f.write(classInfoStr1)
	    f.close()
    with open('conf_classInfo_G8-14.json','w') as f:
        f.write(classInfoStr2)
        f.close()
    print("\nALL DONE !")

def get_merged_cells(sheet):
    return sheet.merged_cells

def get_merged_cells_value(sheet, row_index, col_index):
    merged = get_merged_cells(sheet)
    for (rlow, rhigh, clow, chigh) in merged:
        if (row_index >= rlow and row_index < rhigh):
            if (col_index >= clow and col_index < chigh):
                cell_value = sheet.cell_value(rlow, clow)
                if cell_value:
                    mr = row_index
                    return cell_value
                else:
                    break
    return None


importlib.reload(sys)
main()