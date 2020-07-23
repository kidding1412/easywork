#!/usr/bin/Python2.7
# -*- coding:utf8 -*-
import xlrd
import Tkinter
from Tkinter import *
import tkFileDialog
import ttk
import docx
import re
from docx import Document

# import sys
# reload(sys)
# sys.setdefaultencoding('utf-8')


# excel表格路径
xls_file = r'C:\test.xls'
# word模板路径
docx_file = r'C:\test.docx'


# 表格初始化,输入路径和sheet索引
def init_table(xls_file_root, sheet_index=0):
    sheet_tmp = xlrd.open_workbook(xls_file_root).sheet_by_index(sheet_index)
    return sheet_tmp


# get cell_value
def get_cell_value(sheet_tmp, xrow, xcol):
    return no_dot(sheet_tmp.cell_value(xrow, xcol))


# get sheetname
def get_sheet_name(sheet_tmp):
    return sheet_tmp.name


# get sheet rows number
def get_sheet_rows_number(sheet_tmp):
    return sheet_tmp.nrows


# get sheet cols number
def get_sheet_cols_number(sheet_tmp):
    return sheet_tmp.ncols


# xls列内容写入list
def x_cols_to_list(sheet_tmp, index=0):
    x_cols = sheet_tmp.col_values(index)
    list_tmp = []
    for i in x_cols[1:]:
        list_tmp.append(i)
        print('1')
    return map(no_dot, list_tmp)


# 打印sheet所有collll
def print_all_col(sheet_tmp):
    num_col = sheet_tmp.ncols
    while num_col > 0:
        print x_cols_to_list(sheet_tmp, sheet_tmp.ncols - num_col)
        num_col = num_col - 1
    return


# xls行内容写入list,输入sheet名，index
def x_rows_to_list(sheet_tmp, index=0):
    x_rows = sheet_tmp.row_values(index)
    list_tmp = []
    for i in x_rows[:]:
        list_tmp.append(i)
        # list_tmp.append(unicode(i, "UTF-8"))
    return map(no_dot, list_tmp)


# 打印sheet所有row
def print_all_row(sheet_tmp):
    num_row = sheet_tmp.nrows
    while num_row > 0:
        print x_rows_to_list(sheet_tmp, sheet_tmp.nrows - num_row)
        num_row = num_row - 1
    return


# print测试所有函数
def run_all(sheet_tmp):
    print get_sheet_name(sheet_tmp)
    print get_sheet_rows_number(sheet_tmp)
    print get_sheet_cols_number(sheet_tmp)
    print_all_row(sheet_tmp)
    print_all_col(sheet_tmp)
    print get_cell_value(sheet_tmp, 1, 1)
    return


# def get_cell(sheet_tmp, rows_index, cols_index):
# run_all(sheet)
# 清空treeview函数
def del_button(self):
    global tree
    x = tree.get_children()
    for item in x:
        tree.delete(item)
    return


# 表格路径选择函数
def select_Path(event):
    print '进入路径设置函数'
    path_ = tkFileDialog.askopenfilename(filetypes=[('xls', 'XLS')])
    print path_
    label_table_pwd_show["text"] = path_
    print '路径合法性判断'
    if path_ == "":
        print "表格路径未选择"
        label_table_pwd_show["text"] = "请选择表格路径~"
        return
    xls_file = label_table_pwd_show["text"]
    global sheet  # 全局变量指定excel中的sheet
    sheet = init_table(xls_file, 0)
    # global tree
    del_button(tree)
    if load_excel_to_treeview(sheet) == 0:
        print '表格加载错误'
    else:
        print '表格加载成功'

    print '路径选择函数结束'
    run_all(sheet)
    # 表头初始化
    tabel_head_init(sheet)
    label_names_selected["text"] = '请选择'


# 初始化表头显示，将excel的第一行填写进treeview表头
def tabel_head_init(sheet_tmp):
    print 'join tabel_head_init'
    list_tmp = x_rows_to_list(sheet_tmp)
    # tree.heading('c' + '1', text=list_tmp[0])
    for i in range(0, len(list_tmp)):
        tree.heading('c' + str(i + 1), text=list_tmp[i])
        print str(i + 1) + ':' + list_tmp[i]
    print 'tabel head init over'


# 模板选择函数
def select_docx(event):
    global docx_file
    print 'join select_docx'
    path_ = tkFileDialog.askopenfilename(filetypes=[('doc', 'docx')])
    print path_
    label_docx_pwd_show["text"] = path_
    print '路径合法性判断'
    if path_ == "":
        print "模板路径未选择"
        label_docx_pwd_show["text"] = "请选择表格路径~"
        return
    docx_file = path_
    return


# list小数点过滤函数
def no_dot(x):
    if isinstance(x, float):
        if 5 > x / 1000 > 1:
            return x
        elif str(x)[-2:] == ".0":
            # print '侦测到小数点0'
            return int(x)
        else:
            return x
    else:
        return x


# 全选
def select_all(event):
    print 'join select_all'
    global tree
    t = tree.get_children()
    x = 0
    for i in t:
        x += 1
    if x == 0:
        label_names_selected['text'] = 'empty table'
        return

    for node in tree.get_children():
        tree.selection_add(node)
    global select_list
    select_list = []
    for item in tree.selection():
        select_list.append(tree.item(item, "values")[0])
        select_No_list.append(tree.index(item))
    label_names_selected['text'] = select_list
    return


# 处理生成docx
def gen_docx(n):
    print "join gen_docx"
    print "row index = %s" % n
    global sheet
    global tree
    list_head = []
    list_table = []
    dict = {}
    # list_head = x_rows_to_list(sheet)
    for i in x_rows_to_list(sheet):
        list_head.append(i)
    # list_table = x_rows_to_list(sheet, n+1)
    for n in x_rows_to_list(sheet, n + 1):
        list_table.append(n)
    for i in range(0, len(list_head)):
        dict[list_head[i]] = list_table[i]
    print dict
    print (docx_file)
    doc = Document(docx_file)
    print list_head
    print list_table
    # 正则匹配所有@XXXX
    p = re.compile("^@(.*?)$")
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                # print cell.text
                if re.match(p, cell.text):
                    x = cell.text[1:]
                    print x
                    y = dict[x]
                    # y = y.decode('utf-8')
                    print y , type(y)
                    print type(y)
                    if isinstance(y, int):
                        field_value = str(y)
                    elif isinstance(y, float):
                        field_value = str(y)
                        print field_value
                    elif isinstance(y, long):
                        field_value = str(y)
                    else:
                        field_value = y
                    # cell.text = y
                    cell.text = field_value
                    #print "2"
                    # cell.text.
                    # cell.text = dict[cell.text]
    # doc.save(output_path.get() + "/" + list_table[0] + ".docx")
    a = ("%s/%s.docx") % (output_path.get(), list_table[0])
    print a
    doc.save(a)
    print "end gen_docx"
    return


# 运行函数
def run(event):
    if output_path == '':
        print "select out put path"
        return
    global select_list
    global select_No_list
    global sheet
    print 'join run'
    for i in select_No_list:
        print 'gen_docx(%d)' % i
        gen_docx(i)
    return


# 选择输出路径
def select_output_Path(self):
    path_ = tkFileDialog.askdirectory()
    output_path.set(path_)
    print output_path.get()
    if output_path.get() != '':
        label_out_put_path["text"] = output_path.get()
    return


# 初始化窗口控件
top = Tk()
top.title("EasyWork")
top.geometry('600x600')
top.resizable(width=False, height=False)
# 第一行label
label_table_pwd_text = Label(top, text="表格路径:")
label_table_pwd_text.grid(row=0, column=0, sticky=SW)
label_table_pwd_show = Label(top, text=xls_file, width=50, height=1)
label_table_pwd_show.grid(row=0, column=1)
label_set_table_pwd_button = Button(top, text="表格设置")
label_set_table_pwd_button.grid(row=0, column=2, sticky=SE)
label_set_table_pwd_button.bind("<ButtonRelease-1>", select_Path)  # 绑定选择路径函数
# 第二行label
label_docx_pwd_text = Label(top, text='模板路径:')
label_docx_pwd_text.grid(row=1, column=0, sticky=SW)
label_docx_pwd_show = Label(top, text=docx_file, width=50, height=1)
label_docx_pwd_show.grid(row=1, column=1)
label_set_docx_pwd_button = Button(top, text="模板设置")
label_set_docx_pwd_button.grid(row=1, column=2, sticky=SE)
label_set_docx_pwd_button.bind("<ButtonRelease-1>", select_docx)  # 绑定模板选择函数
# 第三行显示选择结果的label
label_names_selected = Label(top, text='请选择行', wraplength=520)
label_names_selected.place(x=0, y=370, width=600, height=60)
# 第四行label
label_select_all = Button(top, text='全选')
label_select_all.place(x=0, y=430)
label_select_all.bind("<ButtonRelease-1>", select_all)  # 绑定全选函数
label_run = Button(top, text='Run')
label_run.place(x=545, y=430)
label_run.bind("<ButtonRelease-1>", run)
label_out_put_path = Button(top, text='选择输出路径')
label_out_put_path.place(x=55, y=430, width=490)
label_out_put_path.bind("<ButtonRelease-1>", select_output_Path)
# 第五行文字说明
label_state1 = Label(top, text=r'1. 点击"表格设置"选择表格路径，表格为xls或xlsx')
label_state2 = Label(top, text=r'2. 点击"模板设置"选择模板路径，模板为doc或docx')
label_state3 = Label(top, text=r'3. 点击"输出路径"选择文件输出路径')
label_state4 = Label(top, text=r'4. 选择要输出的行，按ctrl可多选')
label_state5 = Label(top, text=r'P.S. 只有word模板中的表格内容会被填写')
label_state6 = Label(top, text=r'       要填写的内容请用"@"+表格列数据的表头做标记，可参照demo')
label_state1.place(x=0, y=460)
label_state2.place(x=0, y=480)
label_state3.place(x=0, y=500)
label_state4.place(x=0, y=520)
label_state5.place(x=0, y=540)
label_state6.place(x=0, y=560)
# frame组件
frame = Frame(top)
frame.place(x=0, y=60, width=600, height=300)
# 滚动条
scrollBar = Scrollbar(frame)
scrollBar.pack(side=Tkinter.RIGHT, fill=Tkinter.Y)
# treeview填充
tree = ttk.Treeview(frame, columns=('c1', 'c2', 'c3', 'c4', 'c5', 'c6', 'c7'), show="headings",
                    yscrollcommand=scrollBar.set)
# 设置每列宽度和对齐方式
tree.column('c1', width=80, anchor='center')
tree.column('c2', width=60, anchor='center')
tree.column('c3', width=60, anchor='center')
tree.column('c4', width=90, anchor='center')
tree.column('c5', width=100, anchor='center')
tree.column('c6', width=90, anchor='center')
tree.column('c7', width=120, anchor='center')
# 设置每列表头标题文本
tree.heading('c1', text='1')
tree.heading('c2', text='2')
tree.heading('c3', text='3')
tree.heading('c4', text='4')
tree.heading('c5', text='5')
tree.heading('c6', text='6')
tree.heading('c7', text='7')
tree["selectmode"] = "extended"  # 允许按ctrl键高亮选中多行
tree.pack(side=Tkinter.LEFT, fill=Tkinter.Y)
# Treeview组件与垂直滚动条结合
scrollBar.config(command=tree.yview)
# 输出路径
output_path = StringVar()


# 定义并绑定Treeview组件的鼠标左键松开事件
def treeviewClick(event):
    print '事件触发'
    # label_table_pwd_show["text"] = ""
    # for item in tree.selection():
    #     # print label_table_pwd_text.info
    #     print tree.item(item, "values")[0]
    #     # label_table_pwd_show["text"] = label_table_pwd_show["text"] + tree.item(item, "values")[6]
    #     label_table_pwd_show["text"] = len(tree.item(item, "values")[6])
    # for item in tree.focus():
    #     print tree.item(item, "values")[0]
    # for item in tree.selection():
    #     label_table_pwd_show["text"] = tree.item(item, "values")[0]
    # print tree.identify_row(event.y)
    # tree.item(tree.identify_row(event.y), values="1")
    # tree.selection_toggle()
    global select_list
    global select_No_list
    select_No_list = []
    select_list = []
    for item in tree.selection():
        select_list.append(tree.item(item, "values")[0])
        select_No_list.append(tree.index(item))
    label_names_selected['text'] = select_list
    # print select_No_list
    return


# 选中的人名list
select_list = []
# 选中的row_index list
select_No_list = []
# 给tree绑定鼠标松开左键事件
tree.bind('<ButtonRelease-1>', treeviewClick)


# 表格加载函数
def load_excel_to_treeview(sheet_tmp):
    print '进入表格加载函数'
    if get_sheet_rows_number(sheet_tmp) < 2:
        return 0
    print '完成空表格判断'
    # global tree
    for i in range(1, get_sheet_rows_number(sheet_tmp)):
        tree.insert('', i, values=x_rows_to_list(sheet_tmp, i))
        print x_rows_to_list(sheet_tmp, i)
    return 1


top.mainloop()
