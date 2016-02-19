#coding=utf-8
#author: David.Lee

from wrExcel import *
import wrExcel
import TestFrameLib
import ReadConfig

rc = ReadConfig.Test_Config(os.getcwd()+"\\Config.ini")
ope = wrExcel.ExcelDriverLibrary()
tfl = TestFrameLib.TestFrameLib(os.getcwd() + rc.caseDir, rc.host)
ope.open_excel(tfl.sFile)

#获取excel文件的sheet名字
sheet_names = ope.get_sheet_names()
assert isinstance(sheet_names, object)
print sheet_names
for sheet_name in sheet_names:
    col =  ope.get_column_count(sheet_name)
    row =  ope.get_row_count(sheet_name)
    #获取接口的名字
    interface = sheet_name
    if col <> 0 or row <> 0:
        print col, row
        values = []
        #遍历当前接口的测试数据
        for i in range(1,row):
            for j in range(1,col-2):
                #将当前用例的测试数据存放在list里面
                parameter = ope.get_row_values(sheet_name, i, includeEmptyCells=True).pop(j)
                assert isinstance(parameter, object)
                values.append(str(int(parameter[1])))
            print values
            #将整条数据拼接成串
            case = ','.join(values)
            #将接口和数据字符串拼接成参数串
            url = '/?*=[[' + interface + ',[' + case + ']]]'
            print 'qa-soul.shinezoneapp.com'+url
            #执行调用
            res_collection= str(tfl.HTTPRequest(url,'','GET'))
            #取服务器返回值中的code值
            code =  tfl.get_result_code(res_collection)

            if ope.Get_CellValue(sheet_name ,i, col-2) == '':
                ope.Modify_cell(sheet_name, i ,col-2, code,ope.OK_STYLE)
            #根据code值判断用例执行的是否成功
            elif ope.Get_CellValue(sheet_name ,i, col-2) == code and ope.Get_CellValue(sheet_name ,i, col-2) != 'case is bad':
                ope.Modify_cell(sheet_name, i ,col-1, 'passed',ope.OK_STYLE)
            elif ope.Get_CellValue(sheet_name ,i, col-2) == 'case is bad':
                ope.Modify_cell(sheet_name, i ,col-1, 'NT',ope.NT_STYLE)
            else:
                ope.Modify_cell(sheet_name, i ,col-1, 'failed',ope.NG_STYLE)
            #清空list，以便存放下一条用例的数据
            if values.__len__() >= col-3:
                values = []
                continue
    else:
        continue