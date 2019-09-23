# coding: utf-8
#!/usr/bin/python

import sys
import json
import xlrd
import importlib

def main():
    # 读取 excel 文件
    workbook = xlrd.open_workbook('MEY4.xlsx')

    headStr = '{\n"classInfo":[\n'
    tailStr = ']\n}'
    itemHeadStr = '{\n'
    itemTailStr = '\n}'
    classInfoStr = ''
    classInfoStr += headStr

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
            #print(weekday)
            sessionlist = []
            module = []
            mtype = []
            teacher = []
            venue = []

            for ri in range(2,19):
#                print(ri)
                if (sheet.cell(ri, ci).value is None or sheet.cell(ri, ci).value == ''):
                    if get_merged_cells_value(sheet, ri, ci):
                        if (get_merged_cells_value(sheet, ri, ci) == '体育'):
                            continue
                        elif (get_merged_cells_value(sheet, ri, ci) == 'No classes shall be arranged at this period'):
                        	continue
                        elif (get_merged_cells_value(sheet, ri, ci) == 'Physical Education'):
                        	continue
                        else:
                            module.append(str(get_merged_cells_value(sheet, ri, ci)))
                            sessionlist.append(str(int(sheet.cell(ri, 0).value)))
                            if (sheet.cell(ri, ci+1).value is None or sheet.cell(ri, ci+1).value == ''):
                                mtype.append(str(get_merged_cells_value(sheet, ri, ci+1)))
                            else:
                                mtype.append(str(sheet.cell(ri, ci+1).value))
                            teacher.append(str(get_merged_cells_value(sheet, ri, ci+2)))
                            venue.append(str(get_merged_cells_value(sheet, ri, ci+3)))
                    else:
                        continue
                else:
                    if (sheet.cell(ri, ci).value == '体育'):
                        continue
                    elif (get_merged_cells_value(sheet, ri, ci) == 'No classes shall be arranged at this period'):
                        continue
                    elif (get_merged_cells_value(sheet, ri, ci) == 'Physical Education'):
                        continue
                    else:
                        module.append(str(sheet.cell(ri, ci).value))
                        sessionlist.append(str(int(sheet.cell(ri, 0).value)))
                        mtype.append(str(sheet.cell(ri, ci+1).value))
                        teacher.append(str(sheet.cell(ri, ci+2).value))
                        venue.append(str(sheet.cell(ri, ci+3).value))
            module = list(map(lambda x: x.replace('\n', ' ').replace('）', ')').replace('，', ',').replace('  ', ' ').replace('XJME3470', 'Vehicle Design and Analysis').replace('XJME3890', 'Individual Engineering Project').replace('XJME3496', 'Thermofluids 3').replace('XJME3900', 'Finite Element Methods of Analysis').replace('XJME3775', 'Additive Manufacturing').replace('形势与政策7 Current Affairs7', '形势与政策7'), module))
            teacher = list(map(lambda x: x.replace('\n', ' ').replace('  ', ' '), teacher))
            mtype = list(map(lambda x: x.replace('\n', ' ').replace('  ', ' '), mtype))
            venue = list(map(lambda x: x.replace('\n', ' ').replace('  ', ' '), venue))
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

def get_merged_cells(sheet):
    return sheet.merged_cells

def get_merged_cells_value(sheet, row_index, col_index):
    merged = get_merged_cells(sheet)
    for (rlow, rhigh, clow, chigh) in merged:
        if (row_index >= rlow and row_index < rhigh):
            if (col_index >= clow and col_index < chigh):
                cell_value = sheet.cell_value(rlow, clow)
                if cell_value:
                    return cell_value
                else:
                    break
    return None
           
importlib.reload(sys)
main()
