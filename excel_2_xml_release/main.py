from tkinter import *
import tkinter as tk
import choiceFiles
import choicePath
import convert
from globalvar import Var
import tkinter.font as tf
import format
import desk


def main():

    # 定义程序中需要用的全局变量或者全局列表,并存入全局字典

    Var.set_value('excelFile_suffix', [])  # 将excelFile_suffix(带路径带后缀的完整信息的文件名列表)存入全局变量
    Var.set_value('excelFile', [])  # 将excelFile(不带后缀名只有文件名的文件名列表)存入全局变量
    Var.set_value('Done_excelFile', [])  # Done_excelFile(已经转换完成的带路径带后缀的完整信息的文件名列表)存入全局变量,确保同一文件启动一次程序不发生重复转换
    Var.set_value('excelFile_Path', [])  # excelFile_Path(每个转换文件的路径列表)存入全局变量,以便把生成文件放在原路径下
    Var.set_value('excelFile_convert', [])  # excelFile_convert(本次需要转换文件的路径列表)存入全局变量

    '''创建窗口/按钮/提示字符'''

    # 设置弹窗标题和窗口大小
    win = Tk()  # 创建一个Tkinter.Tk()实例
    Var.set_value('win', win)  # 将win存入全局变量,以便在其他py文件调用配置
    # win.withdraw()  # 隐藏窗口
    win.wm_attributes('-topmost', 1)  # 1:置顶窗口  2:取消置顶
    win.title('excel_2_xml_version0.3(Debug Version: excel_2_xml_06)')  # 窗口标题
    win.geometry('1000x400+450+200')  # 窗口大小
    win.resizable(height=False)  # 规定窗口不可缩放

    # 创建下拉菜单File，然后将其加入到顶级的菜单栏中
    bar = tk.Menu(win)  # 创建菜单栏
    menu = tk.Menu(bar, tearoff=0)
    menu.add_command(label="Choice a dir", command=choicePath.ChoicePath().choicePath)
    menu.add_command(label="Choice a file", command=choiceFiles.ChoiceFiles().choiceFiles)
    menu.add_separator()
    menu.add_command(label="clear excelFile", command=desk.Win().Clear_list)
    menu.add_command(label="clear debugInfo", command=desk.Win().Clear_info)
    menu.add_separator()
    menu.add_command(label="Exit", command=win.quit)
    bar.add_cascade(label="Option", menu=menu)

    menu_01 = tk.Menu(bar, tearoff=False)
    user_choice = IntVar()
    user_choice.set('Matrix')  # 默认选择第一个
    menu_01.add_radiobutton(label='Matrix', variable=user_choice, value='Matrix', command=format.Format().Matrix)
    menu_01.add_radiobutton(label='BDCM event', variable=user_choice, value='BDCM event', command=format.Format().Event)
    menu_01.add_radiobutton(label='BDCM cycle', variable=user_choice, value='BDCM cycle', command=format.Format().Cycle)
    bar.add_cascade(label="Format", menu=menu_01)
    Var.set_value('Format', 'Matrix')
    print(Var.get_value('Format'))
    # 显示v的值
    format_ver = tk.Label(win, textvariable=user_choice)
    Var.set_value('Format', 'Matrix')
    format_ver.grid(row=5, column=0, sticky=E)
    # 创建一个list界面
    # listbox与滚动条相结合
    lv = StringVar([], [])
    listbox = Listbox(win, listvariable=lv, width=140, height=12)  # 创建listbox
    Var.set_value('listbox', listbox)  # 将select_01存入全局变量,以便在其他py文件调用配置
    listbox.grid(row=2, column=0, sticky=E + W)
    scrollbar2 = Scrollbar(win)  # 创建滚动条
    scrollbar2.grid(row=2, column=1, sticky=N + S + W + E)
    # 配置listbox和滚动条
    listbox.config(yscrollcommand=scrollbar2.set)
    scrollbar2.config(command=listbox.yview)

    info = tk.Label(listbox, text='excel file', width=13, borderwidth=1, relief="solid")
    info.place(relx=1.0, rely=1.0, x=-2, y=-2, anchor="se")

    # 创建一个text界面
    text = Text(win, width=140, height=12, bd=3, relief='groove')
    Var.set_value('text', text)  # 将select_01存入全局变量,以便在其他py文件调用配置
    text.grid(row=3, column=0, padx=0, pady=2)
    # 创建滚动条
    scrollbar = Scrollbar(win)
    scrollbar.grid(row=3, column=1, sticky=N + S + W + E)
    # 配置文本和滚动条
    text.config(yscrollcommand=scrollbar.set, state=DISABLED)
    scrollbar.config(command=text.yview)

    info_01 = tk.Label(text, text='convert info', width=13, borderwidth=1, relief="solid")
    info_01.place(relx=1.0, rely=1.0, x=-2, y=-2, anchor="se")

    # # 选择特定文件的按钮
    # # 配置选择文件按钮,并传给变量select_01
    # select_01 = Button(win, text='Choice files', command=choiceFiles.ChoiceFiles().choiceFiles)
    # Var.set_value('select_01', select_01)  # 将select_01存入全局变量,以便在其他py文件调用配置
    # select_01.grid(row=5, column=0, pady=10)  # 打包select_01按钮,设置按钮的一些显示属性

    # # 选择特定文件夹的按钮
    # # 配置选择文件路径按钮,并传给变量select_02
    # select_02 = Button(win, text='Choice path', command=choicePath.ChoicePath().choicePath)
    # Var.set_value('select_02', select_02)  # 将select_02存入全局变量,以便在其他py文件调用配置
    # select_02.grid(row=6, column=0, pady=10)  # 打包select_02按钮,设置按钮的一些显示属性

    # 确定转换按钮
    # 配置确认转换按钮,并传给变量convert
    a = Button(win, text='convert', width=10, height=1, command=convert.Convert().convert)
    Var.set_value('convert', a)
    a.place(relx=0.45, rely=0.9)  # 打包select_02按钮,设置按钮的一些显示属性

    # # 创建显示此时正在转换的文件的文本
    # State = Label(win)
    # Var.set_value('State', State)  # 将State存入全局变量,以便在其他py文件调用配置

    # 创建显示转换进度文本
    State_01 = Label(win)
    Var.set_value('State_01', State_01)  # 将State存入全局变量,以便在其他py文件调用配置

    listbox_row = 0
    Var.set_value('listbox_row', listbox_row)  # 将State存入全局变量,以便在其他py文件调用配置

    # 设置提示信息的字体
    ft = tf.Font(family='宋体', size=10)  # 有很多参数
    text.tag_config('tag', foreground='green', font=ft)  # 设置tag即插入文字的大小,颜色等
    text.tag_config('tag_01', foreground='red', font=ft)  # 设置tag即插入文字的大小,颜色等
    text.tag_config('tag_02', foreground='blue', font=ft)  # 设置tag即插入文字的大小,颜色等

    # # 创建在窗口显示所添加的文件的文件名的文本
    # excel_name = Label(win)
    # Var.set_value('excel_name', excel_name)  # 将State存入全局变量,以便在其他py文件调用配置

    # text.config(state=NORMAL)
    # text.insert('insert', user_choice, 'tag_01')  # 申明使用tag中的设置
    # text.yview(MOVETO, 1.0)
    # text.config(state=DISABLED)

    # 显示菜单
    win.config(menu=bar)
    mainloop()
    # win.mainloop()  # 显示设置好的窗口和按钮


if __name__ == '__main__':
    main()
