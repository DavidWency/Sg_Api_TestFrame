#coding:utf=8
__author__ = 'ZZ'


import os
import natsort
from operator import itemgetter
from datetime import datetime,timedelta
from xlrd import open_workbook, cellname, xldate_as_tuple,error_text_from_code
import xlwt
from xlutils.copy import copy
from version import Version

_version_ = Version

class ExcelDriverLibrary:
    def __init__(self):
        self.wb = None
        self.rb = None
        self.Pid = None
        self.rowIndex = None
        self.sheetNum = None
        self.SheetNames = None
        self.fileName = None
        self.OK_STYLE = xlwt.easyxf('font:color-index blue,bold on;')
        self.NG_STYLE = xlwt.easyxf('font:color-index red,bold on;')
        self.NT_STYLE = xlwt.easyxf('font:color-index yellow,bold on;')

    def open_excel(self,filename):
        tempDir = 'D:\\Codes\\PycharmProjects\\ExcelDataDriver\\ResFile\\'
        # filename = 'TestFile.xls'
        try:
            if filename.find(':')==1:
                self.rb = open_workbook(filename,formatting_info=True)
            else:
                print 'Opening file at %s' % filename
                #self.wb = open_workbook(os.path.join("/",self.tmpDir,filename),formatting_info=True, on_demand=True)
                self.rb = open_workbook(os.path.join(tempDir,filename),formatting_info=True, on_demand=True)
                filename = os.path.join(tempDir,filename)
            self.fileName = filename
            self.SheetNames = self.rb.sheet_names
        except Exception,e:
            print str(e)
    def get_sheet_names(self):
        SheetNames = self.rb.sheet_names()
        return SheetNames

    def get_number_of_sheets(self):
        sheetNum = self.rb.nsheets
        return sheetNum

    #得到当前sheet的列数
    def get_column_count(self,SheetName):
        sheet = self.rb.sheet_by_name(SheetName)
        return sheet.ncols
    #得到当前sheet的行数
    def get_row_count(self,SheetName):
        sheet = self.rb.sheet_by_name(SheetName)
        return sheet.nrows

    def get_column_values(self,SheetName,column,includeEmptyCells=True):
        sheet = self.rb.sheet_by_name(SheetName)
        data = {}
        for row_index in range(sheet.nrows):
            cell = cellname(row_index,int(column))
            value = sheet.cell(row_index,int(column)).value
            data[cell] = int(value)
        if includeEmptyCells is True:
            sortedData = natsort.natsorted(data.items(),key=itemgetter(0))
            return sortedData
        else:
            data = dict([(k,v) for (k,v) in data.items() if v])
            OrderedData = natsort.natsorted(data.items(),key=itemgetter(0))
            return OrderedData

    def get_row_values(self,SheetName,row,includeEmptyCells=True):
        sheet = self.rb.sheet_by_name(SheetName)
        data = {}
        for col_index in range(sheet.ncols):
            cell = cellname(int(row),col_index)
            value = sheet.cell(int(row),col_index).value
            data[cell] = value
        if includeEmptyCells is True:
            sortedData = natsort.natsorted(data.items(),key=itemgetter(0))
            return sortedData
        else:
            data = dict([(k,v) for (k,v) in data.items() if v])
            OrderedData = natsort.natsorted(data.items(),key=itemgetter(0))
            return OrderedData

    def get_sheet_values(self,SheetName,includeEmptyCells=True):
        sheet = self.rb.sheet_by_name(SheetName)
        data = {}
        for row_index in range(sheet.nrows):
            for col_index in range(sheet.ncols):
                cell = cellname(row_index,col_index)
                value = sheet.cell(row_index,col_index).value
                data[cell] = value
        if includeEmptyCells is True:
            sortedData = natsort.natsorted(data.items(),key=itemgetter(0))
            return sortedData
        else:
            data = dict([(k,v) for (k,v) in data.items() if v])
            OrderedData = natsort.natsorted(data.items(),key=itemgetter(0))
            return OrderedData

    def get_workBoot_value(self,includeEmptyCells=True):

        sheetData = []
        workbookData = []
        for sheet_name in self.SheetNames:
            if includeEmptyCells is True:
                sheetData = self.get_sheet_values(sheet_name)
            else:
                sheetData = self.get_sheet_values(sheet_name,False)
            sheetData.insert(0,sheet_name)
            workbookData.append(sheetData)
        return workbookData

    def read_cell_data_by_name(self,SheetName,cell_name):

        # Uses the cell name to return the data from that cell.
        sheet = self.rb.sheet_by_name(SheetName)
        for row_index in range(sheet.nrows):
            for col_index in range(sheet.ncols):
                cell = cellname(row_index,col_index)
                if cell_name ==cell:
                    cellValue = sheet.cell(row_index,col_index).value
        return cellValue

    def read_cell_data_by_coordinates(self, SheetName, column, row):
        # Uses the column and row to return the data from that cell.
        my_sheet_index = self.SheetNames.index(SheetName)
        sheet = self.rb.sheet_by_index(my_sheet_index)
        cellValue = sheet.cell(int(row), int(column)).value
        return cellValue

    def Get_Sheet_Index(self,SheetName):
        sheetNum = self.get_number_of_sheets()
        for sheet_index in range(sheetNum):
            Current_SheetName =  self.rb.sheet_names()[sheet_index]
            if SheetName == Current_SheetName.encode('gb2312').decode('gb2312'):
                return sheet_index

    def Get_Cell_Data_By_PID(self,SheetName,Pid,CellName):
        sheet = self.rb.sheet_by_name(SheetName)
        for row_index in range(sheet.nrows):
            Cell_Value = sheet.cell(int(row_index),0).value
            if Pid == Cell_Value:
                break
        for col_index in range(sheet.ncols):
            Cell_TitValue = sheet.cell(0,int(col_index)).value
            if CellName == Cell_TitValue:
                break
        cellValue = sheet.cell(row_index, col_index).value
        return cellValue



    def Modify_index_cell(self,SheetName,Pid,CellName,CellValue):
        CellName_index = self._get_CellName_index(CellName,SheetName)
        Pid_index = self._get_PID_index(Pid,SheetName)
        rb = open_workbook(self.fileName,formatting_info=True)
        try:
            sheet_index = rb._sheet_names.index(SheetName)
        except Exception,e:
            print u'请输入正确的Sheet名,如果是中文,记得在变量名前面加u.具体错误消息为：',e
            exit()
        wb =copy(rb)
        sheet = wb.get_sheet(sheet_index)

        sheet.write(Pid_index,CellName_index,CellValue)
        wb.save(self.fileName)

    def Modify_cell(self, sheetName ,row ,col ,cellValue, cellStyle):
        global sheet_index
        rd = open_workbook(self.fileName,formatting_info=True)
        try:
            sheet_index = self.Get_Sheet_Index(sheetName)
        except Exception,e:
            print u'请输入正确的Sheet名,如果是中文,记得在变量名前面加u.具体错误消息为：',e
            exit()
        cp = copy(rd)
        sheet = cp.get_sheet(sheet_index)
        sheet.write(row, col, cellValue, cellStyle)
        cp.save(self.fileName)

    def Get_CellValue(self, sheetName, row, col):
        sheet = self.rb.sheet_by_name(sheetName)
        cellValue = sheet.cell(row, col).value
        return cellValue
