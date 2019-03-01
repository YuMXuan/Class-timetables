# coding: utf-8
#!/usr/bin/python

import sys
import json
import xlrd
import importlib

def main():
    # 读取 excel 文件
    workbook = xlrd.open_workbook('1819Y3ME.xlsx')

    headStr = '{\n"classInfo":[\n'
    tailStr = ']\n}'
    itemHeadStr = '{\n'
    itemTailStr = '\n}'
    classInfoStr = ''
    classInfoStr += headStr

#    merge = []

    week = 0

    for s in range(0,len(workbook.sheets())):
        week = week+1
        print(week*5)
        sheet = workbook.sheet_by_index(s)
        numOfRow = sheet.nrows
        numOfCol = sheet.ncols
        weekday = 0
        for ci in range(1,21,4):
            weekday = weekday+1
            sessionlist = []
            module = []
            mtype = []
            teacher = []
            venue = []

            for ri in range(2,19):          
                module.append(str(sheet.cell(ri, ci).value))
                while '' in module:
                    module.remove('')
                    break
                else:
                    sessionlist.append(str(int(sheet.cell(ri, 0).value)))
                    mtype.append(str(sheet.cell(ri, ci+1).value))
                    teacher.append(str(sheet.cell(ri, ci+2).value))
                    venue.append(str(sheet.cell(ri, ci+3).value))              
            i = 0
            for className in module:
                itemClassInfoStr = ""
                itemClassInfoStr += itemHeadStr + '"className":"' + className + '",\n'
                itemClassInfoStr += '"week":' + str(week) + ',\n'
                itemClassInfoStr += '"weekday":' + str(weekday) + ',\n'
                itemClassInfoStr += '"session":' + sessionlist[i] + ',\n'
                itemClassInfoStr += '"mtype":"' + mtype[i] + '",\n'
                itemClassInfoStr += '"teacher":"' + teacher[i] + '",\n'
                itemClassInfoStr += '"venue":"' + venue[i] + '"'
                itemClassInfoStr += itemTailStr
                itemClassInfoStr += ","
                classInfoStr += itemClassInfoStr
                i += 1
    classInfoStr = classInfoStr.strip(",")
    classInfoStr += tailStr
    with open('conf_classInfo.json','w') as f:

	    f.write(classInfoStr)
	    f.close()
    print("\nALL DONE !")
           
importlib.reload(sys)
main()
