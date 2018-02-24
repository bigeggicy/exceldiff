# -*- coding: utf-8 -*-

import wx
import wx.xrc
import wx.grid
import wx.dataview
import os
import xlrd

row_result = []
col_result = []
cell_result = []
excel_old_row_labels = []
excel_new_row_labels = []
excel_old_col_labels = []
excel_new_col_labels = []

translate ={"pass":u"无修改","delete":u"删除","insert":u"新增"}


def convert_to_title(num):
    num += 1
    capitals = [chr(x) for x in range(ord('A'), ord('Z') + 1)]
    result = []
    while num > 0:
        result.append(capitals[(num - 1) % 26])
        num = (num - 1) // 26
    result.reverse()
    return ''.join(result)


def find_nth_element(list,target,n):
    i=0
    for i in range(0,len(list)):
        if target is list[i]:
            n-=1
        if n==0:
            return i
    if i==len(list):
        return -1


def nmMatrix(n, m, v):
    r = []
    for i in range(n):
        mr = []
        for j in range(m):
            mr.append(v)
        r.append(mr)
    return r


# 单元格比较函数
def cell_compare(cell1, cell2):
    if cell1.value==cell2.value:
        return True
    return False


# 行列的函数比较，行/列所有单元格不同才返回不同
def row_compare(row1, row2):
    n = min(len(row1), len(row2))
    for i in range(0, n):
        if (row1[i].value == ""and row2[i].value =="") is False and row1[i].value == row2[i].value:
            return True  # 只要有一个相同，返回相同，视为相同行/列
    all_empty_flap = 1
    for i in range(0,n):
        if (row1[i].value == ""and row2[i].value =="") is False:
            all_empty_flap = 0
    if all_empty_flap == 1:
        return True
    return False


def dfs(list1, list2, f, v, n, m):
    if f[n][m] >= 0:
        return f[n][m]
    if n == 0 and m == 0:
        f[n][m] = 0
        return 0
    if n == 0:
        v[n][m] = 2
        f[n][m] = dfs(list1, list2, f, v, n, m - 1) + 1
        return f[n][m]
    if m == 0:
        v[n][m] = 1
        f[n][m] = dfs(list1, list2, f, v, n - 1, m) + 1
        return f[n][m]
    if row_compare(list1[n - 1], list2[m - 1]):
        f[n][m] = dfs(list1, list2, f, v, n - 1, m - 1)
        return f[n][m]
    if dfs(list1, list2, f, v, n, m - 1) > dfs(list1, list2, f, v, n - 1, m):
        f[n][m] = dfs(list1, list2, f, v, n - 1, m) + 1
        v[n][m] = 1
    else:
        f[n][m] = dfs(list1, list2, f, v, n, m - 1) + 1
        v[n][m] = 2
    return f[n][m]


def row_find_operation(n, m, v):
    if n == 0 and m == 0:
        return
    if v[n][m] == 0:
        row_find_operation(n - 1, m - 1, v)
        row_result.append("pass")
    elif v[n][m] == 1:
        row_find_operation(n - 1, m, v)
        row_result.append("delete")
    else:
        row_find_operation(n, m - 1, v)
        row_result.append("insert")


def col_find_operation(n, m, v):
    if n == 0 and m == 0:
        return
    if v[n][m] == 0:
        col_find_operation(n - 1, m - 1, v)
        col_result.append("pass")
    elif v[n][m] == 1:
        col_find_operation(n - 1, m, v)
        col_result.append("delete")
    else:
        col_find_operation(n, m - 1, v)
        col_result.append("insert")


def row_func(list1, list2):
    n = len(list1)
    m = len(list2)
    del row_result[:]
    f = nmMatrix(n + 1, m + 1, -1)
    v = nmMatrix(n + 1, m + 1, 0)
    dfs(list1, list2, f, v, n, m)
    row_find_operation(n, m, v)


def col_func(list1, list2):
    n = len(list1)
    m = len(list2)
    del col_result[:]
    f = nmMatrix(n + 1, m + 1, -1)
    v = nmMatrix(n + 1, m + 1, 0)
    dfs(list1, list2, f, v, n, m)
    col_find_operation(n, m, v)


class ExcelDiff(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"excel_diff", pos=wx.DefaultPosition,
                          size=wx.Size(1400, 700), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)
        self.old_data = None
        self.new_data = None
        self.old_file_path=""
        self.new_file_path=""
        self.SetSizeHintsSz(wx.DefaultSize, wx.DefaultSize)

        sbSizer1 = wx.StaticBoxSizer(wx.StaticBox(self, wx.ID_ANY, u"比较新旧excel"), wx.VERTICAL)

        self.upload_old = wx.Button(sbSizer1.GetStaticBox(), wx.ID_ANY, u"上传旧Excel表", wx.DefaultPosition,
                                    wx.DefaultSize, 0)
        self.Bind(wx.EVT_BUTTON, self.upload_old_file, self.upload_old)

        sbSizer1.Add(self.upload_old, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.upload_new = wx.Button(sbSizer1.GetStaticBox(), wx.ID_ANY, u"上传新Excel表", wx.DefaultPosition,
                                    wx.DefaultSize, 0)
        sbSizer1.Add(self.upload_new, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)
        self.Bind(wx.EVT_BUTTON, self.upload_new_file, self.upload_new)

        self.sheet_compare_button=wx.Button(sbSizer1.GetStaticBox(), wx.ID_ANY, u"比较sheet差异", wx.DefaultPosition, wx.DefaultSize, 0)
        sbSizer1.Add(self.sheet_compare_button, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)
        self.Bind(wx.EVT_BUTTON, self.sheet_compare, self.sheet_compare_button)

        gSizer1 = wx.GridSizer(0, 2, 0, 0)
        self.sheet_compare_choice = []
        self.sheet_select = wx.Choice(sbSizer1.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
                                      self.sheet_compare_choice, 0)
        self.sheet_select.SetSelection(0)
        gSizer1.Add(self.sheet_select, 0, wx.ALL | wx.ALIGN_RIGHT, 5)

        self.compare_button = wx.Button(sbSizer1.GetStaticBox(), wx.ID_ANY, u"比较选中sheet的差异", wx.DefaultPosition, wx.DefaultSize, 0)
        sbSizer1.Add(self.compare_button, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)
        self.Bind(wx.EVT_BUTTON, self.compare_file, self.compare_button)

        gSizer1.Add(self.compare_button, 0, wx.ALL | wx.ALIGN_RIGHT, 5)
        sbSizer1.Add(gSizer1, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        fgSizer3 = wx.FlexGridSizer(0, 2, 0, 0)
        fgSizer3.SetFlexibleDirection(wx.HORIZONTAL)
        fgSizer3.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        sbSizer6 = wx.StaticBoxSizer(wx.StaticBox(sbSizer1.GetStaticBox(), wx.ID_ANY, u"旧Excel表"), wx.VERTICAL)

        self.excel_old = wx.grid.Grid(sbSizer6.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)

        # Grid
        self.excel_old.CreateGrid(5, 7)
        self.excel_old.EnableEditing(False)
        self.excel_old.EnableGridLines(True)
        self.excel_old.EnableDragGridSize(False)
        self.excel_old.SetMargins(0, 0)

        # Columns
        self.excel_old.EnableDragColMove(False)
        self.excel_old.EnableDragColSize(False)
        self.excel_old.SetColLabelSize(30)
        self.excel_old.SetColLabelAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)

        # Rows
        self.excel_old.EnableDragRowSize(False)
        self.excel_old.SetRowLabelSize(80)
        self.excel_old.SetRowLabelAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)

        # Label Appearance

        # Cell Defaults
        self.excel_old.SetDefaultCellAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
        sbSizer6.Add(self.excel_old, 0, wx.ALL, 5)

        fgSizer3.Add(sbSizer6, 1, wx.EXPAND, 5)

        sbSizer7 = wx.StaticBoxSizer(wx.StaticBox(sbSizer1.GetStaticBox(), wx.ID_ANY, u"新Excel表"), wx.VERTICAL)

        self.excel_new = wx.grid.Grid(sbSizer7.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)

        # Grid
        self.excel_new.CreateGrid(5, 7)
        self.excel_new.EnableEditing(False)
        self.excel_new.EnableGridLines(True)
        self.excel_new.EnableDragGridSize(False)
        self.excel_new.SetMargins(0, 0)

        # Columns
        self.excel_new.EnableDragColMove(False)
        self.excel_new.EnableDragColSize(False)
        self.excel_new.SetColLabelSize(30)
        self.excel_new.SetColLabelAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)

        # Rows
        self.excel_new.EnableDragRowSize(False)
        self.excel_new.SetRowLabelSize(80)
        self.excel_new.SetRowLabelAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)

        # Label Appearance

        # Cell Defaults
        self.excel_new.SetDefaultCellAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
        sbSizer7.Add(self.excel_new, 0, wx.ALL, 5)

        fgSizer3.Add(sbSizer7, 1, wx.EXPAND, 5)

        sbSizer1.Add(fgSizer3, 1, wx.EXPAND, 5)

        fgSizer4 = wx.FlexGridSizer(0, 4, 0, 0)
        fgSizer4.SetFlexibleDirection(wx.HORIZONTAL)
        fgSizer4.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        sbSizer8 = wx.StaticBoxSizer(wx.StaticBox(sbSizer1.GetStaticBox(), wx.ID_ANY, u"行增删"), wx.VERTICAL)

        self.row_compare_text = wx.StaticText(sbSizer8.GetStaticBox(), wx.ID_ANY, u"共计新增 行，删除 行", wx.DefaultPosition,
                                              wx.DefaultSize, 0)
        self.row_compare_text.Wrap(-1)
        sbSizer8.Add(self.row_compare_text, 0, wx.ALL, 5)

        self.row_compare = wx.grid.Grid(sbSizer8.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)

        # Grid
        self.row_compare.CreateGrid(5, 2)
        self.row_compare.SetColLabelValue(0, u"改动")
        self.row_compare.SetColLabelValue(1, u"行号")
        self.row_compare.EnableEditing(False)
        self.row_compare.EnableGridLines(True)
        self.row_compare.EnableDragGridSize(False)
        self.row_compare.SetMargins(0, 0)

        # Columns
        self.row_compare.EnableDragColMove(False)
        self.row_compare.EnableDragColSize(False)
        self.row_compare.SetColLabelSize(30)
        self.row_compare.SetColLabelAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)

        # Rows
        self.row_compare.EnableDragRowSize(False)
        self.row_compare.SetRowLabelSize(0)
        self.row_compare.SetRowLabelAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)

        # Label Appearance
        # self.row_compare.SetLabelBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_HIGHLIGHT))

        # Cell Defaults
        self.row_compare.SetDefaultCellAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)
        sbSizer8.Add(self.row_compare, 0, wx.ALL, 5)

        fgSizer4.Add(sbSizer8, 1, wx.EXPAND, 5)

        sbSizer10 = wx.StaticBoxSizer(wx.StaticBox(sbSizer1.GetStaticBox(), wx.ID_ANY, u"列增删"), wx.VERTICAL)

        self.col_compare_text = wx.StaticText(sbSizer10.GetStaticBox(), wx.ID_ANY, u"共计新增 行，删除 行", wx.DefaultPosition,
                                              wx.DefaultSize, 0)
        self.col_compare_text.Wrap(-1)
        sbSizer10.Add(self.col_compare_text, 0, wx.ALL, 5)

        self.col_compare = wx.grid.Grid(sbSizer10.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)

        # Grid
        self.col_compare.CreateGrid(5, 2)
        self.col_compare.SetColLabelValue(0, u"改动")
        self.col_compare.SetColLabelValue(1, u"列号")
        self.col_compare.EnableEditing(False)
        self.col_compare.EnableGridLines(True)
        self.col_compare.EnableDragGridSize(False)
        self.col_compare.SetMargins(0, 0)

        # Columns
        self.col_compare.EnableDragColMove(False)
        self.col_compare.EnableDragColSize(False)
        self.col_compare.SetColLabelSize(30)
        self.col_compare.SetColLabelAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)

        # Rows
        self.col_compare.EnableDragRowSize(False)
        self.col_compare.HideRowLabels()
        self.col_compare.SetRowLabelAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)

        # Label Appearance

        # Cell Defaults
        self.col_compare.SetDefaultCellAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
        sbSizer10.Add(self.col_compare, 0, wx.ALL, 5)

        fgSizer4.Add(sbSizer10, 1, wx.EXPAND, 5)

        sbSizer12 = wx.StaticBoxSizer(wx.StaticBox(sbSizer1.GetStaticBox(), wx.ID_ANY, u"单元格修改"), wx.VERTICAL)

        self.cell_compare_text = wx.StaticText(sbSizer12.GetStaticBox(), wx.ID_ANY, u"共计改动 个单元格", wx.DefaultPosition,
                                              wx.DefaultSize, 0)
        self.cell_compare_text.Wrap(-1)
        sbSizer12.Add(self.cell_compare_text, 0, wx.ALL, 5)

        self.cell_compare = wx.grid.Grid(sbSizer12.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)

        # Grid
        self.cell_compare.CreateGrid(5, 3)
        self.cell_compare.SetColLabelValue(0, u"坐标")
        self.cell_compare.SetColLabelValue(1, u"旧值")
        self.cell_compare.SetColLabelValue(2, u"新值")
        self.cell_compare.EnableEditing(False)
        self.cell_compare.EnableGridLines(True)
        self.cell_compare.EnableDragGridSize(False)
        self.cell_compare.SetMargins(0, 0)

        # Columns
        self.cell_compare.EnableDragColMove(False)
        self.cell_compare.EnableDragColSize(False)
        self.cell_compare.SetColLabelSize(30)
        self.cell_compare.SetColLabelAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)

        # Rows
        self.cell_compare.EnableDragRowSize(False)
        self.cell_compare.HideRowLabels()
        self.cell_compare.SetRowLabelAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)

        # Label Appearance

        # Cell Defaults
        self.cell_compare.SetDefaultCellAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
        sbSizer12.Add(self.cell_compare, 0, wx.ALL, 5)

        fgSizer4.Add(sbSizer12, 1, wx.EXPAND, 5)

        sbSizer13 = wx.StaticBoxSizer(wx.StaticBox(sbSizer1.GetStaticBox(), wx.ID_ANY, u"sheet增删"), wx.VERTICAL)

        self.sheet_compare_text = wx.StaticText(sbSizer13.GetStaticBox(), wx.ID_ANY, u"共计增加 个sheet，删除 个sheet", wx.DefaultPosition,
                                               wx.DefaultSize, 0)
        self.sheet_compare_text.Wrap(-1)
        sbSizer13.Add(self.sheet_compare_text, 0, wx.ALL, 5)

        self.sheet_compare_result = wx.grid.Grid(sbSizer13.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)

        # Grid
        self.sheet_compare_result.CreateGrid(5, 2)
        self.sheet_compare_result.SetColLabelValue(0, u"改动")
        self.sheet_compare_result.SetColLabelValue(1, u"sheet名")
        self.sheet_compare_result.EnableEditing(False)
        self.sheet_compare_result.EnableGridLines(True)
        self.sheet_compare_result.EnableDragGridSize(False)
        self.sheet_compare_result.SetMargins(0, 0)

        # Columns
        self.sheet_compare_result.EnableDragColMove(False)
        self.sheet_compare_result.EnableDragColSize(False)
        self.sheet_compare_result.SetColLabelSize(30)
        self.sheet_compare_result.SetColLabelAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)

        # Rows
        self.sheet_compare_result.EnableDragRowSize(False)
        self.sheet_compare_result.HideRowLabels()
        self.sheet_compare_result.SetRowLabelAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)

        # Label Appearance

        # Cell Defaults
        self.sheet_compare_result.SetDefaultCellAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)
        sbSizer13.Add(self.sheet_compare_result, 0, wx.ALL, 5)

        fgSizer4.Add(sbSizer13, 1, wx.EXPAND, 5)

        sbSizer1.Add(fgSizer4, 1, wx.EXPAND, 5)

        self.SetSizer(sbSizer1)
        self.Layout()
        self.Centre(wx.BOTH)

    def __del__(self):
        pass

    def upload_old_file(self, e):
        wildcard = "Excel Files (*.xlsx)|*.xlsx"
        dlg = wx.FileDialog(self, "Choose a file", os.getcwd(), "", wildcard, wx.OPEN)

        if dlg.ShowModal() == wx.ID_OK:
            self.old_file_path = dlg.GetPath()
            self.old_data = xlrd.open_workbook(self.old_file_path)
        dlg.Destroy()
        self.sheet_select.Clear()

    def upload_new_file(self,e):
        wildcard = "Excel Files (*.xlsx)|*.xlsx"
        dlg = wx.FileDialog(self, "Choose a file", os.getcwd(), "", wildcard, wx.OPEN)

        if dlg.ShowModal() == wx.ID_OK:
            self.new_file_path=dlg.GetPath()
            self.new_data = xlrd.open_workbook(self.new_file_path)
        dlg.Destroy()
        self.sheet_select.Clear()

    def sheet_compare(self,e):
        if self.old_data is None:
            self.error(u"请上传旧Excel文件")
            return
        if self.new_data is None:
            self.error(u"请上传新Excel文件")
            return
        if self.old_file_path==self.new_file_path:
            self.error(u"新旧文件是同一文件")
            return

        self.sheet_compare_result.ClearGrid()
        old_sheets_set = set(self.old_data.sheet_names())
        new_sheets_set = set(self.new_data.sheet_names())
        same_sheets = list(old_sheets_set.intersection(new_sheets_set))
        self.sheet_select.SetItems(same_sheets)
        delete_sheets = old_sheets_set.difference(new_sheets_set)
        insert_sheets = new_sheets_set.difference(old_sheets_set)
        counter = 0
        for i in delete_sheets:
            self.sheet_compare_result.SetCellValue(counter,0,u"删除")
            self.sheet_compare_result.SetCellValue(counter, 1, i)
            counter += 1
        for i in insert_sheets:
            self.sheet_compare_result.SetCellValue(counter, 0, u"新增")
            self.sheet_compare_result.SetCellValue(counter, 1, i)
            counter += 1
        self.sheet_compare_text.SetLabelText(u"共计新增"+str(len(insert_sheets))+u"个sheet，删除"+
                                             str(len(delete_sheets))+u"个sheet")
        self.Layout()

    def compare_file(self, e):
        self.excel_old.ClearGrid()
        self.excel_new.ClearGrid()
        self.excel_old.SetBackgroundColour(wx.NullColour)
        self.excel_old.ForceRefresh()
        self.excel_new.SetBackgroundColour(wx.NullColour)
        self.excel_new.ForceRefresh()
        self.row_compare.ClearGrid()
        self.col_compare.ClearGrid()
        self.cell_compare.ClearGrid()

        # if self.old_data is None:
        #     self.error(u"请上传旧Excel文件")
        #     return
        # if self.new_data is None:
        #     self.error(u"请上传新Excel文件")
        #     return
        # if self.old_file_path==self.new_file_path:
        #     self.error(u"新旧文件是同一文件")
        #     return
        # old_sheets = self.old_data.sheet_names()
        # new_sheets = self.new_data.sheet_names()
        # # 找出第一个符合条件的同名sheet，old_table和new_table读取同名sheet内容
        # target_sheet_name = None
        # for x in old_sheets:
        #     if x in new_sheets:
        #         target_sheet_name = x
        #         break
        # if target_sheet_name is None:
        #     self.error(u"不存在同名sheet")
        #     return
        target_sheet_name = self.sheet_select.GetStringSelection()
        if target_sheet_name is u"":
            self.error(u"请选择正确的同名sheet")
            return
        old_table = self.old_data.sheet_by_name(target_sheet_name)
        new_table = self.new_data.sheet_by_name(target_sheet_name)

        # 行列差异的基本思路：把行/列一整行视为对象，表格视为有序的列对象的组成,DP找listA转成listB的最小操作
        # 单元格差异思路：把新旧表格涉及行列差异的单元格删除，再进行单元格的比较即可
        #
        # 计算行列差别，操作保存在全局变量result中
        #
        old_rows = []
        new_rows = []
        old_cols = []
        new_cols = []
        for i in range(0, old_table.nrows):
            old_rows.append(old_table.row(i))
        for i in range(0, new_table.nrows):
            new_rows.append(new_table.row(i))
        row_func(old_rows, new_rows)  # DP
        for i in range(0, old_table.ncols):
            old_cols.append(old_table.col(i))
        for i in range(0, new_table.ncols):
            new_cols.append(new_table.col(i))
        col_func(old_cols, new_cols)
        # 单元格比较
        i = 0
        m = 0
        # i,j为旧表格的行列，m、n为新表格行列
        del cell_result[:]
        while i < old_table.nrows and m < new_table.nrows:
            while row_result[i] == "delete":
                i += 1
                if i >= old_table.nrows:
                    break
            while row_result[m] == "insert":
                m += 1
                if m >= new_table.nrows:
                    break
            if i >= old_table.nrows or m >= new_table.nrows:
                break
            else:
                j = 0
                n = 0
                while j < old_table.ncols and n < new_table.ncols:
                    while col_result[j] == "delete":
                        j += 1
                        if j >= old_table.ncols:
                            break
                    while col_result[n] == "insert":
                        n += 1
                        if n >= new_table.ncols:
                            break
                    if j >= old_table.ncols or n >= new_table.ncols:
                        break
                    if not cell_compare(old_table.cell(i, j), new_table.cell(m, n)):
                        cell_result.append([i, j, m, n, old_table.cell(i, j).value, new_table.cell(m, n).value])
                    j += 1
                    n += 1
            i += 1
            m += 1
        # 数据处理完毕
        #
        # 显示结果
        self.row_compare_text.SetLabelText(
            u"共计新增" + str(row_result.count("insert")) + u"行,删除" + str(row_result.count("delete")) + u"行")
        self.col_compare_text.SetLabelText(
            u"共计新增" + str(col_result.count("insert")) + u"列,删除" + str(col_result.count("delete")) + u"列")
        self.cell_compare_text.SetLabelText(
            u"共计改动" + str(len(cell_result)) + u"个单元格")
        if row_result.count("delete") + row_result.count("insert") - self.row_compare.GetNumberRows() > 0:
            self.row_compare.AppendRows(row_result.count("delete") + row_result.count("insert") - self.row_compare.GetNumberRows())
        tempt = 0
        insert_counter = 0
        delete_counter = 0
        row_insert_result = []
        row_delete_result = []
        for i in range(0, len(row_result)):
            if row_result[i] is "delete":
                self.row_compare.SetCellValue(tempt, 0, translate[row_result[i]])
                self.row_compare.SetCellValue(tempt, 1, str(i + 1-insert_counter))
                row_delete_result.append(i + 1-insert_counter)
                tempt += 1
                delete_counter += 1
            if row_result[i] is "insert":
                self.row_compare.SetCellValue(tempt, 0, translate[row_result[i]])
                self.row_compare.SetCellValue(tempt, 1, str(i + 1 - delete_counter))
                row_insert_result.append(i + 1 - delete_counter)
                tempt += 1
                insert_counter += 1
        if col_result.count("delete") + col_result.count("insert") - self.col_compare.GetNumberRows() > 0:
            self.col_compare.AppendRows(col_result.count("delete") + col_result.count("insert") - self.col_compare.GetNumberRows())
        tempt = 0
        delete_counter = 0
        insert_counter = 0
        col_insert_result = []
        col_delete_result = []
        for i in range(0, len(col_result)):
            if col_result[i] is "delete":
                self.col_compare.SetCellValue(tempt, 0, translate[col_result[i]])
                self.col_compare.SetCellValue(tempt, 1, convert_to_title(i-insert_counter))
                col_delete_result.append(convert_to_title(i-insert_counter))
                tempt += 1
                delete_counter += 1
            if col_result[i] is "insert":
                self.col_compare.SetCellValue(tempt, 0, translate[col_result[i]])
                self.col_compare.SetCellValue(tempt, 1, convert_to_title(i - delete_counter))
                col_insert_result.append(convert_to_title(i - delete_counter))
                tempt += 1
                insert_counter += 1
        tempt = 0
        if len(cell_result) > self.cell_compare.GetNumberRows():
            self.cell_compare.AppendRows(len(cell_result) - self.cell_compare.GetNumberRows())
        for i in cell_result:
            self.cell_compare.SetCellValue(tempt, 0, "[" + str(i[0] + 1) + "," + convert_to_title(i[1]) + "]" + ",[" + str(
                i[2] + 1) + "," + convert_to_title(i[3]) + "]")

            if type(i[4]) == int or type(i[4]) == float:
                self.cell_compare.SetCellValue(tempt, 1, str(i[4]))
            else:
                self.cell_compare.SetCellValue(tempt, 1, i[4])
            # if type(i[4]) == unicode:
            #     self.cell_compare.SetCellValue(tempt, 1,i[4])
            # else:
            #     self.cell_compare.SetCellValue(tempt, 1, str(i[4]))

            if type(i[5]) == int or type(i[5]) == float:
                self.cell_compare.SetCellValue(tempt, 2, str(i[5]))
            else:
                self.cell_compare.SetCellValue(tempt, 2, i[5])
            tempt += 1
        # 结果输出结束
        #########################################
        # 绘制表格old_excel和new_excel
        # 扩充行列,默认(5,7）,根据结果扩充
        if old_table.nrows+row_result.count("insert")-self.excel_old.GetNumberRows() > 0:
            self.excel_old.AppendRows(old_table.nrows+row_result.count("insert")-self.excel_old.GetNumberRows())
        elif old_table.nrows+row_result.count("insert")-self.excel_old.GetNumberRows() < 0:
            self.excel_old.DeleteRows(numRows=self.excel_old.GetNumberRows()-old_table.nrows - row_result.count("insert"))
        # 备忘，deleterows不同appendrow，delete与insert对应多一个参数
        if old_table.ncols+col_result.count("insert")-self.excel_old.GetNumberCols() > 0:
            self.excel_old.AppendCols(old_table.ncols+col_result.count("insert")-self.excel_old.GetNumberCols())
        elif old_table.ncols+col_result.count("insert")-self.excel_old.GetNumberCols() < 0:
            self.excel_old.DeleteCols(numCols=self.excel_old.GetNumberCols() - old_table.ncols - col_result.count("insert"))

        if new_table.nrows+row_result.count("delete")-self.excel_new.GetNumberRows() > 0:
            self.excel_new.AppendRows(new_table.nrows + row_result.count("delete") - self.excel_new.GetNumberRows())
        elif new_table.nrows+row_result.count("delete")-self.excel_new.GetNumberRows() < 0:
            self.excel_new.DeleteRows(numRows=self.excel_new.GetNumberRows()-new_table.nrows - row_result.count("delete"))

        if new_table.ncols+col_result.count("delete")-self.excel_new.GetNumberCols() > 0:
            self.excel_new.AppendCols(new_table.ncols+col_result.count("delete")-self.excel_new.GetNumberCols())
        elif new_table.ncols+col_result.count("delete")-self.excel_new.GetNumberCols() < 0:
            self.excel_new.DeleteCols(numCols=self.excel_new.GetNumberCols() - new_table.ncols - col_result.count("delete"))

        # 制作新旧表格表头的list
        del excel_old_row_labels[:]
        del excel_new_row_labels[:]
        del excel_old_col_labels[:]
        del excel_new_col_labels[:]
        for i in range(0, self.excel_old.GetNumberRows()):
            if i >= len(row_result) or row_result[i] != "insert":
                excel_old_row_labels.append(i+1)
            else:
                i += 1
                excel_old_row_labels.append("")
        for i in range(0, self.excel_old.GetNumberCols()):
            if i >= len(col_result) or col_result[i] != "insert":
                excel_old_col_labels.append(convert_to_title(i))
            else:
                excel_old_col_labels.insert(i-1, "")
                i += 1
        for i in range(0,self.excel_new.GetNumberRows()):
            excel_new_row_labels.append(i+1)
        for i in range(0, self.excel_new.GetNumberCols()):
            excel_new_col_labels.append(convert_to_title(i))

        for i in row_delete_result:
            excel_new_row_labels.insert(excel_old_row_labels.index(i), "")
        for i in col_delete_result:
            excel_new_col_labels.insert(excel_new_col_labels.index(i)+1,"")
        # 绘制表头
        for i in range(0,self.excel_old.GetNumberRows()):
            self.excel_old.SetRowLabelValue(i,str(excel_old_row_labels[i]))
        for i in range(0,self.excel_old.GetNumberCols()):
            self.excel_old.SetColLabelValue(i,str(excel_old_col_labels[i]))
        for i in range(0,self.excel_new.GetNumberRows()):
            self.excel_new.SetRowLabelValue(i,str(excel_new_row_labels[i]))
        for i in range(0,self.excel_new.GetNumberCols()):
            self.excel_new.SetColLabelValue(i,str(excel_new_col_labels[i]))
        # 开始制作绘画old_excel
        m = 0
        for i in range(0, old_table.nrows):
            while len(self.excel_old.GetRowLabelValue(m)) == 0:
                m += 1
            n = 0
            for j in range(0, old_table.ncols):
                while len(self.excel_old.GetColLabelValue(n)) is 0:
                    n += 1
                if type(old_table.cell_value(i, j)) is int or type(old_table.cell_value(i, j)) is float:
                    self.excel_old.SetCellValue(m, n, str(old_table.cell_value(i, j)))
                else:
                    self.excel_old.SetCellValue(m, n, old_table.cell_value(i, j))
                n += 1
                if n >= self.excel_old.GetNumberCols():
                    break
            m += 1
            if m >= self.excel_old.GetNumberRows():
                break
        # 绘制新表格new_excel
        m = 0
        for i in range(0, new_table.nrows):
            while len(self.excel_new.GetRowLabelValue(m)) == 0:
                m += 1
            n = 0
            for j in range(0, new_table.ncols):
                while len(self.excel_new.GetColLabelValue(n)) is 0:
                    n += 1
                if type(new_table.cell_value(i, j)) == int or type(new_table.cell_value(i, j)) == float:
                    self.excel_new.SetCellValue(m, n, str(new_table.cell_value(i, j)))
                else:
                    self.excel_new.SetCellValue(m, n, new_table.cell_value(i, j))
                n += 1
                if n >= self.excel_new.GetNumberCols():
                    break
            m += 1
            if m >= self.excel_new.GetNumberRows():
                break
        # 上色
        for i in range(0,self.excel_old.GetNumberRows()):
            for j in range(0, self.excel_old.GetNumberCols()):
                self.excel_old.SetCellBackgroundColour(i, j, wx.NullColour)
        for i in range(0,self.excel_new.GetNumberRows()):
            for j in range(0, self.excel_new.GetNumberCols()):
                self.excel_new.SetCellBackgroundColour(i, j, wx.NullColour)
        for i in row_delete_result:
            index= excel_old_row_labels.index(i)
            for j in range(0,self.excel_old.GetNumberCols()):
                self.excel_old.SetCellBackgroundColour(index, j, '#F08080')
            for j in range(0,self.excel_new.GetNumberCols()):
                self.excel_new.SetCellBackgroundColour(index, j, '#F08080')
        for i in row_insert_result:
            index=excel_new_row_labels.index(i)
            for j in range(0, self.excel_old.GetNumberCols()):
                self.excel_old.SetCellBackgroundColour(index, j, '#B0C4DE')
            for j in range(0, self.excel_new.GetNumberCols()):
                self.excel_new.SetCellBackgroundColour(index, j, '#B0C4DE')
        for i in col_delete_result:
            index= excel_old_col_labels.index(i)
            for j in range(0,self.excel_old.GetNumberRows()):
                self.excel_old.SetCellBackgroundColour(j,index,'#F08080')
            for j in range(0,self.excel_new.GetNumberRows()):
                self.excel_new.SetCellBackgroundColour(j, index, '#F08080')
        for i in col_insert_result:
            index = excel_new_col_labels.index(i)
            for j in range(0, self.excel_old.GetNumberRows()):
                self.excel_old.SetCellBackgroundColour(j,index, '#B0C4DE')
            for j in range(0, self.excel_new.GetNumberRows()):
                self.excel_new.SetCellBackgroundColour(j,index, '#B0C4DE')
        for i in cell_result:
            self.excel_old.SetCellBackgroundColour(excel_old_row_labels.index(i[0]+1),excel_old_col_labels.index(convert_to_title(i[1])),
                                                   "#FFFF00")
            self.excel_new.SetCellBackgroundColour(excel_new_row_labels.index(i[2]+1),
                                                   excel_new_col_labels.index(convert_to_title(i[3])),
                                                   "#FFFF00")
        self.excel_old.ForceRefresh()
        self.excel_new.ForceRefresh()
        self.Layout()
        # 联动
        self.Bind(wx.grid.EVT_GRID_CELL_LEFT_CLICK, self.row_result_onclick, self.row_compare)
        self.Bind(wx.grid.EVT_GRID_CELL_LEFT_CLICK, self.col_result_onclick, self.col_compare)
        self.Bind(wx.grid.EVT_GRID_CELL_LEFT_CLICK, self.cell_result_onclick, self.cell_compare)

    def row_result_onclick(self,e):
        self.excel_old.ClearSelection()
        self.excel_new.ClearSelection()

        if self.row_compare.GetCellValue(e.GetRow(), 0) == u"新增":
            new_insert_row = int(self.row_compare.GetCellValue(e.GetRow(), 1))
            self.excel_new.SelectRow(excel_new_row_labels.index(new_insert_row))
            self.excel_new.GoToCell(excel_new_row_labels.index(new_insert_row),0)
            insert_count=0
            for i in range(0,e.GetRow()):
                if self.row_compare.GetCellValue(i,0) == u"新增":
                    insert_count += 1
            old_insert_row = find_nth_element(excel_old_row_labels, "", insert_count+1)
            self.excel_old.SelectRow(old_insert_row)
            self.excel_old.GoToCell(old_insert_row,0)
        if self.row_compare.GetCellValue(e.GetRow(),0) == u"删除":
            old_delete_row = int(self.row_compare.GetCellValue(e.GetRow(),1))
            self.excel_old.SelectRow(excel_old_row_labels.index(old_delete_row))
            self.excel_old.GoToCell(excel_old_row_labels.index(old_delete_row),0)
            delete_count = 0
            for i in range(0, e.GetRow()):
                if self.row_compare.GetCellValue(i,0) == u"删除":
                    delete_count += 1
            new_delete_row = find_nth_element(excel_new_row_labels, "", delete_count + 1)
            self.excel_new.SelectRow(new_delete_row)
            self.excel_new.GoToCell(new_delete_row,0)
        e.Skip()

    def col_result_onclick(self,e):
        self.excel_old.ClearSelection()
        self.excel_new.ClearSelection()
        if self.col_compare.GetCellValue(e.GetRow(),0) == u"新增":
            new_insert_col = self.col_compare.GetCellValue(e.GetRow(),1)
            self.excel_new.SelectCol(excel_new_col_labels.index(new_insert_col))
            self.excel_new.GoToCell(0,excel_new_col_labels.index(new_insert_col))
            insert_count = 0
            for i in range(0, e.GetRow()):
                if self.col_compare.GetCellValue(i,0) == u"新增":
                    insert_count += 1
            old_insert_col = find_nth_element(excel_old_col_labels, "", insert_count + 1)
            self.excel_old.SelectCol(old_insert_col)
            self.excel_old.GoToCell(0,old_insert_col)
        if self.col_compare.GetCellValue(e.GetRow(),0) == u"删除":
            old_delete_col = self.col_compare.GetCellValue(e.GetRow(),1)
            self.excel_old.SelectCol(excel_old_col_labels.index(old_delete_col))
            self.excel_old.GoToCell(0,excel_old_col_labels.index(old_delete_col))
            delete_count = 0
            for i in range(0, e.GetRow()):
                if self.col_compare.GetCellValue(i,0) == u"删除":
                    delete_count += 1
            new_delete_col = find_nth_element(excel_new_col_labels, "", delete_count + 1)
            self.excel_new.SelectCol(new_delete_col)
            self.excel_new.GoToCell(0,new_delete_col)
        e.Skip()

    def cell_result_onclick(self,e):
        if e.GetRow() >= len(cell_result):
            e.Skip()
            return
        i = cell_result[e.GetRow()]

        old_cell_row = excel_old_row_labels.index(i[0] + 1)
        old_cell_col = excel_old_col_labels.index(convert_to_title(i[1]))
        new_cell_row = excel_new_row_labels.index(i[2] + 1)
        new_cell_col = excel_new_col_labels.index(convert_to_title(i[3]))

        self.excel_old.ClearSelection()
        self.excel_old.SelectBlock(old_cell_row, old_cell_col,old_cell_row, old_cell_col)
        self.excel_old.GoToCell(old_cell_row, old_cell_col)
        self.excel_old.MakeCellVisible(old_cell_row, old_cell_col)
        self.excel_new.ClearSelection()
        self.excel_new.SelectBlock(new_cell_row, new_cell_col,new_cell_row, new_cell_col)
        self.excel_new.MakeCellVisible(new_cell_row, new_cell_col)
        self.excel_new.GoToCell(new_cell_row, new_cell_col)
        e.Skip()

    # 错误提示框
    def error(self, message):
        wx.MessageBox(message, "Error", wx.OK | wx.ICON_INFORMATION)


class App(wx.App):
    def OnInit(self):
        self.frame = ExcelDiff(parent=None)
        self.frame.Show()
        self.SetTopWindow(self.frame)
        return True


if __name__ == '__main__':
    app = App()
    app.MainLoop()

