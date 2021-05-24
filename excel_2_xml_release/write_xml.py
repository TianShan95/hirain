from xml.dom import minidom  # 创建dom文件模块
import datetime  # 用于生成文件的名称里带上日期
import xlrd  # 读取excel模块
from tkinter import *
from globalvar import Var  # 导入全局字典

global message_node


class Excel2Xml(object):

    # 传入表格，正在转换第几个文件，这个文件的路径数组，这个文件的不带后缀的文件名数组，转换完成文件名数组,带后缀的文件名数组
    def __init__(self, excel, num, excel_path, excel_name, done_excel, excel_name_suffix):
        self.excel = excel
        self.Num = num
        self.excelFile_Path = excel_path
        self.excelFile = excel_name
        # self.Done_excelFile = done_excel
        self.excelFile_suffix = excel_name_suffix

        self.result = ''

    def read_excel(self):

        text = Var.get_value('text')  # 在字典中取出之前存入的text
        try:
            data = xlrd.open_workbook(self.excel)  # 打开文件，文件内容存到data
            Format = Var.get_value('Format')
            table = data.sheet_by_name(Format)  # 打开第一个名字为Matrix的sheet
            if Format == 'Matrix':
                self.write_xml_Matrix(table)
                self.result = 'success'
            if Format == 'BDCM event':
                self.write_xml_Event(table)
                self.result = 'success'
            if Format == 'BDCM cycle':
                self.write_xml_Cycle(table)
                self.result = 'success'
            return self.result
        except Exception as err:
            print('错误信息：{0}'.format(err))
            text.config(state=NORMAL)
            text.insert('insert', '>>>' + str(self.Num + 1) + ' : err info：{0}'.format(err) + '\n', 'tag_02')
            text.yview(MOVETO, 1.0)
            text.config(state=DISABLED)
            self.result = 'fail'
            return self.result

    def write_xml_Matrix(self, sheet):  # 传入进来excel里的一个sheet

        print('进了 write_xml_Matrix')
        # 1.创建DOM树对象
        dom = minidom.Document()
        # 2.创建根节点。每次都要用DOM对象来创建任何节点。
        config_node = dom.createElement('config')
        # 3.用DOM对象添加根节点
        dom.appendChild(config_node)
        # 4.设置该节点的属性
        config_node.setAttribute('order', 'intel')

        # 用DOM对象创建元素子节点channel
        channel_node = dom.createElement('channel')
        # 用父节点对象添加元素子节点
        config_node.appendChild(channel_node)
        # 设置该节点的属性
        channel_node.setAttribute('index', '0')

        first_time = 1
        global message_node
        IF_same_signal_name = []
        text = Var.get_value('text')  # 在字典中取出之前存入的text
        for rowNum in range(sheet.nrows):  # 一行一行的遍历
            rvalue = sheet.row_values(rowNum)  # 这一行数据给rvalue
            if rowNum > 1:  # 去掉表头，从第二行开始处理
                if sheet.cell(rowNum, 0).value != '' and first_time == 1:  # 如果是第一次写入xml，因为不是第一次，所以不需要给上个节点添加文本节点filter

                    message_name = rvalue[0]
                    message_id = rvalue[1]
                    message_node = dom.createElement("message")
                    channel_node.appendChild(message_node)
                    message_node.setAttribute('id', message_id)
                    message_node.setAttribute('name', message_name)
                    first_time = 0
                elif sheet.cell(rowNum, 0).value != '' and first_time == 0:  # 进来这里就是写完message里所有的signal
                    filter_node = dom.createElement('filter')  # 上个节点添加上文本节点filter
                    message_node.appendChild(filter_node)
                    # 也用DOM创建文本节点，把文本节点（文字内容）看成子节点
                    filter_text = dom.createTextNode('normal')
                    # 用添加了文本的节点对象（看成文本节点的父节点）添加文本节点
                    filter_node.appendChild(filter_text)

                    message_name = rvalue[0]
                    message_id = rvalue[1]
                    message_node = dom.createElement("message")
                    channel_node.appendChild(message_node)
                    message_node.setAttribute('id', message_id)
                    message_node.setAttribute('name', message_name)
                else:
                    signal_name = rvalue[6]
                    startbit = int(rvalue[13])
                    bitlen = int(rvalue[14])
                    offset = int(rvalue[16])
                    factor = int(rvalue[15])
                    signed = rvalue[22]

                    # 用DOM对象创建元素子节点message
                    signal_node = dom.createElement('signal')
                    # 用父节点对象添加元素子节点
                    message_node.appendChild(signal_node)

                    # 设置该节点的属性
                    signal_node.setAttribute('name', signal_name)
                    signal_node.setAttribute('startbit', str(startbit))
                    signal_node.setAttribute('bitlen', str(bitlen))
                    signal_node.setAttribute('offset', str(offset))
                    signal_node.setAttribute('factor', str(factor))
                    signal_node.setAttribute('signed', str(signed))

                    # 把signal_name和已存入列表中的signal_name进行比对，如果发现有相同的signal_name在界面中打印出文件名和signal_name
                    for i in IF_same_signal_name:
                        if i == signal_name:
                            text.config(state=NORMAL)
                            text.insert('insert',
                                        '>>>' + str(self.Num + 1) + ' : ' + self.excel + ' ' + signal_name + ' repeat\n', 'tag')
                            text.yview(MOVETO, 1.0)
                            text.config(state=DISABLED)
                    IF_same_signal_name.append(signal_name)

        # 把最后一个message添加上文本节点filter
        filter_node = dom.createElement('filter')
        message_node.appendChild(filter_node)
        # 也用DOM创建文本节点，把文本节点（文字内容）看成子节点
        filter_text = dom.createTextNode('normal')
        # 用添加了文本的节点对象（看成文本节点的父节点）添加文本节点
        filter_node.appendChild(filter_text)

        nowTime = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')  # 现在

        try:  # 把生成的文件再放入一开始选择的文件夹里
            with open(self.excelFile_Path[self.Num] + '/' + self.excelFile[self.Num] + '_' + nowTime + r'_Matrix.xml',
                      'w', encoding='UTF-8') as fh:
                # writexml()第一个参数是目标文件对象，第二个参数是根节点的缩进格式，第三个参数是其他子节点的缩进格式，
                # 第四个参数制定了换行格式，第五个参数制定了xml内容的编码。
                dom.writexml(fh, indent='', addindent='  ', newl='\n', encoding='UTF-8')
                print('写入xml OK!')  # 写入成功的标志
                # self.Done_excelFile.append(self.excelFile_suffix[self.Num])  # 把转换成功的文件存起来，避免下次重复转换
        except Exception as err:
            print('错误信息：{0}'.format(err))

    def write_xml_Event(self, sheet):  # 传入进来excel里的一个sheet

        print('进了 write_xml_Event')
        # 1.创建DOM树对象
        dom = minidom.Document()
        # 2.创建根节点。每次都要用DOM对象来创建任何节点。
        config_node = dom.createElement('config')
        # 3.用DOM对象添加根节点
        dom.appendChild(config_node)
        # 4.设置该节点的属性
        config_node.setAttribute('order', 'intel')

        # 用DOM对象创建元素子节点channel
        channel_node = dom.createElement('channel')
        # 用父节点对象添加元素子节点
        config_node.appendChild(channel_node)
        # 设置该节点的属性
        channel_node.setAttribute('index', '0')

        first_time = 1
        global message_node
        IF_same_signal_name = []
        text = Var.get_value('text')  # 在字典中取出之前存入的text
        print(sheet.nrows)
        for rowNum in range(sheet.nrows):  # 一行一行的遍历
            rvalue = sheet.row_values(rowNum)  # 这一行数据给rvalue
            if rowNum > 0:  # 去掉表头，从第二行开始处理
                if sheet.cell(rowNum, 0).value != '' and first_time == 1:  # 如果是第一次写入xml，因为不是第一次，所以不需要给上个节点添加文本节点filter
                    message_id = rvalue[0]
                    message_node = dom.createElement("message")
                    channel_node.appendChild(message_node)
                    message_node.setAttribute('id', message_id)

                    signal_id = int(rvalue[2])
                    print('rvalue[2]')
                    print(rvalue[2])
                    print('signal_id')
                    print(signal_id)
                    print(str(signal_id))
                    signal_name = rvalue[3]
                    bitlen = int(rvalue[5])
                    bytelen = int(rvalue[6])
                    factor = rvalue[7]

                    # 用DOM对象创建元素子节点message
                    signal_node = dom.createElement('signal')
                    # 用父节点对象添加元素子节点
                    message_node.appendChild(signal_node)

                    # 设置该节点的属性
                    signal_node.setAttribute('name', signal_name)
                    signal_node.setAttribute('signal_id', str(signal_id))
                    signal_node.setAttribute('bitlen', str(bitlen))
                    signal_node.setAttribute('bytelen', str(bytelen))
                    signal_node.setAttribute('factor', str(factor))

                    first_time = 0
                elif sheet.cell(rowNum, 0).value != '' and first_time == 0:  # 进来这里就是写完message里所有的signal
                    filter_node = dom.createElement('filter')  # 上个节点添加上文本节点filter
                    message_node.appendChild(filter_node)
                    # 也用DOM创建文本节点，把文本节点（文字内容）看成子节点
                    filter_text = dom.createTextNode('normal')
                    # 用添加了文本的节点对象（看成文本节点的父节点）添加文本节点
                    filter_node.appendChild(filter_text)

                    message_id = rvalue[0]
                    message_node = dom.createElement("message")
                    channel_node.appendChild(message_node)
                    message_node.setAttribute('id', message_id)

                    signal_id = rvalue[2]
                    signal_name = rvalue[3]
                    bitlen = int(rvalue[5])
                    bytelen = int(rvalue[6])
                    factor = rvalue[7]

                    # 用DOM对象创建元素子节点message
                    signal_node = dom.createElement('signal')
                    # 用父节点对象添加元素子节点
                    message_node.appendChild(signal_node)

                    # 设置该节点的属性
                    signal_node.setAttribute('name', signal_name)
                    signal_node.setAttribute('signal_id', signal_id)
                    signal_node.setAttribute('bitlen', str(bitlen))
                    signal_node.setAttribute('bytelen', str(bytelen))
                    signal_node.setAttribute('factor', str(factor))

                    # 把signal_name和已存入列表中的signal_name进行比对，如果发现有相同的signal_name在界面中打印出文件名和signal_name
                    for i in IF_same_signal_name:
                        if i == signal_name:
                            text.config(state=NORMAL)
                            text.insert('insert',
                                        '>>>' + str(self.Num + 1) + ' : ' + self.excel + ' ' + signal_name + ' repeat\n', 'tag')
                            text.yview(MOVETO, 1.0)
                            text.config(state=DISABLED)
                    IF_same_signal_name.append(signal_name)

                elif rvalue[2] != '':
                    signal_id = rvalue[2]
                    signal_name = rvalue[3]
                    bitlen = int(rvalue[5])
                    bytelen = int(rvalue[6])
                    factor = rvalue[7]

                    # 用DOM对象创建元素子节点message
                    signal_node = dom.createElement('signal')
                    # 用父节点对象添加元素子节点
                    message_node.appendChild(signal_node)

                    # 设置该节点的属性
                    signal_node.setAttribute('name', signal_name)
                    signal_node.setAttribute('signal_id', signal_id)
                    signal_node.setAttribute('bitlen', str(bitlen))
                    signal_node.setAttribute('bytelen', str(bytelen))
                    signal_node.setAttribute('factor', str(factor))

                    # 把signal_name和已存入列表中的signal_name进行比对，如果发现有相同的signal_name在界面中打印出文件名和signal_name
                    for i in IF_same_signal_name:
                        if i == signal_name:
                            text.config(state=NORMAL)
                            text.insert('insert',
                                        '>>>' + str(self.Num + 1) + ' : ' + self.excel + ' ' + signal_name + ' repeat\n', 'tag')
                            text.yview(MOVETO, 1.0)
                            text.config(state=DISABLED)
                    IF_same_signal_name.append(signal_name)

        # 把最后一个message添加上文本节点filter
        filter_node = dom.createElement('filter')
        message_node.appendChild(filter_node)
        # 也用DOM创建文本节点，把文本节点（文字内容）看成子节点
        filter_text = dom.createTextNode('normal')
        # 用添加了文本的节点对象（看成文本节点的父节点）添加文本节点
        filter_node.appendChild(filter_text)

        nowTime = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')  # 现在

        try:  # 把生成的文件再放入一开始选择的文件夹里
            with open(self.excelFile_Path[self.Num] + '/' + self.excelFile[self.Num] + '_' +
                      nowTime + r'_BDCM event.xml',
                      'w', encoding='UTF-8') as fh:
                # writexml()第一个参数是目标文件对象，第二个参数是根节点的缩进格式，第三个参数是其他子节点的缩进格式，
                # 第四个参数制定了换行格式，第五个参数制定了xml内容的编码。
                dom.writexml(fh, indent='', addindent='  ', newl='\n', encoding='UTF-8')
                print('写入xml OK!')  # 写入成功的标志
                # self.Done_excelFile.append(self.excelFile_suffix[self.Num])  # 把转换成功的文件存起来，避免下次重复转换
        except Exception as err:
            print('错误信息：{0}'.format(err))

    def write_xml_Cycle(self, sheet):  # 传入进来excel里的一个sheet

        print('进了 write_xml_Cycle')
        # 1.创建DOM树对象
        dom = minidom.Document()
        # 2.创建根节点。每次都要用DOM对象来创建任何节点。
        config_node = dom.createElement('config')
        # 3.用DOM对象添加根节点
        dom.appendChild(config_node)
        # 4.设置该节点的属性
        config_node.setAttribute('order', 'intel')

        # 用DOM对象创建元素子节点channel
        channel_node = dom.createElement('channel')
        # 用父节点对象添加元素子节点
        config_node.appendChild(channel_node)
        # 设置该节点的属性
        channel_node.setAttribute('index', '0')

        first_time = 1
        global message_node
        IF_same_signal_name = []
        text = Var.get_value('text')  # 在字典中取出之前存入的text
        print(sheet.nrows)
        for rowNum in range(sheet.nrows):  # 一行一行的遍历
            rvalue = sheet.row_values(rowNum)  # 这一行数据给rvalue
            # print(rvalue[0] + rvalue[1])
            if rowNum > 0:  # 去掉表头，从第二行开始处理
                if sheet.cell(rowNum, 0).value != '' and first_time == 1:  # 如果是第一次写入xml，因为不是第一次，所以不需要给上个节点添加文本节点filter
                    message_id = rvalue[0]
                    message_node = dom.createElement("message")
                    channel_node.appendChild(message_node)
                    message_node.setAttribute('id', message_id)

                    signal_name = rvalue[1]
                    bitlen = int(rvalue[3])
                    bytelen = int((bitlen + 7) / 8)
                    factor = rvalue[5]

                    # 用DOM对象创建元素子节点message
                    signal_node = dom.createElement('signal')
                    # 用父节点对象添加元素子节点
                    message_node.appendChild(signal_node)

                    # 设置该节点的属性
                    signal_node.setAttribute('name', signal_name)
                    signal_node.setAttribute('bitlen', str(bitlen))
                    signal_node.setAttribute('bytelen', str(bytelen))
                    signal_node.setAttribute('factor', str(factor))

                    first_time = 0
                elif sheet.cell(rowNum, 0).value != '' and first_time == 0:  # 进来这里就是写完message里所有的signal
                    filter_node = dom.createElement('filter')  # 上个节点添加上文本节点filter
                    message_node.appendChild(filter_node)
                    # 也用DOM创建文本节点，把文本节点（文字内容）看成子节点
                    filter_text = dom.createTextNode('normal')
                    # 用添加了文本的节点对象（看成文本节点的父节点）添加文本节点
                    filter_node.appendChild(filter_text)

                    message_id = rvalue[0]
                    message_node = dom.createElement("message")
                    channel_node.appendChild(message_node)
                    message_node.setAttribute('id', message_id)

                    signal_name = rvalue[1]
                    bitlen = int(rvalue[3])
                    bytelen = int((bitlen + 7) / 8)
                    factor = rvalue[5]

                    # 用DOM对象创建元素子节点message
                    signal_node = dom.createElement('signal')
                    # 用父节点对象添加元素子节点
                    message_node.appendChild(signal_node)

                    # 设置该节点的属性
                    signal_node.setAttribute('name', signal_name)
                    signal_node.setAttribute('bitlen', str(bitlen))
                    signal_node.setAttribute('bytelen', str(bytelen))
                    signal_node.setAttribute('factor', str(factor))

                    # 把signal_name和已存入列表中的signal_name进行比对，如果发现有相同的signal_name在界面中打印出文件名和signal_name
                    for i in IF_same_signal_name:
                        if i == signal_name:
                            text.config(state=NORMAL)
                            text.insert('insert',
                                        '>>>' + str(self.Num + 1) + ' : ' + self.excel + ' ' + signal_name + ' repeat\n', 'tag')
                            text.yview(MOVETO, 1.0)
                            text.config(state=DISABLED)
                    IF_same_signal_name.append(signal_name)
                elif rvalue[2] != '':
                    signal_name = rvalue[1]
                    bitlen = int(rvalue[3])
                    bytelen = int((bitlen + 7) / 8)
                    factor = rvalue[5]

                    # 用DOM对象创建元素子节点message
                    signal_node = dom.createElement('signal')
                    # 用父节点对象添加元素子节点
                    message_node.appendChild(signal_node)

                    # 设置该节点的属性
                    signal_node.setAttribute('name', signal_name)
                    signal_node.setAttribute('bitlen', str(bitlen))
                    signal_node.setAttribute('bytelen', str(bytelen))
                    signal_node.setAttribute('factor', str(factor))

                    # 把signal_name和已存入列表中的signal_name进行比对，如果发现有相同的signal_name在界面中打印出文件名和signal_name
                    for i in IF_same_signal_name:
                        if i == signal_name:
                            text.config(state=NORMAL)
                            text.insert('insert',
                                        '>>>' + str(self.Num + 1) + ' : ' + self.excel + ' ' + signal_name + ' repeat\n', 'tag')
                            text.yview(MOVETO, 1.0)
                            text.config(state=DISABLED)
                    IF_same_signal_name.append(signal_name)

        # 把最后一个message添加上文本节点filter
        filter_node = dom.createElement('filter')
        message_node.appendChild(filter_node)
        # 也用DOM创建文本节点，把文本节点（文字内容）看成子节点
        filter_text = dom.createTextNode('normal')
        # 用添加了文本的节点对象（看成文本节点的父节点）添加文本节点
        filter_node.appendChild(filter_text)

        nowTime = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')  # 现在

        try:  # 把生成的文件再放入一开始选择的文件夹里
            with open(self.excelFile_Path[self.Num] + '/' + self.excelFile[self.Num] + '_' +
                      nowTime + r'_BDCM cycle.xml',
                      'w', encoding='UTF-8') as fh:
                # writexml()第一个参数是目标文件对象，第二个参数是根节点的缩进格式，第三个参数是其他子节点的缩进格式，
                # 第四个参数制定了换行格式，第五个参数制定了xml内容的编码。
                dom.writexml(fh, indent='', addindent='  ', newl='\n', encoding='UTF-8')
                print('写入xml OK!')  # 写入成功的标志
                # self.Done_excelFile.append(self.excelFile_suffix[self.Num])  # 把转换成功的文件存起来，避免下次重复转换
        except Exception as err:
            print('错误信息：{0}'.format(err))
