#!/usr/bin/env python
# -*- coding: utf-8 -*-
import io
import re
from datetime import datetime
import csv
import copy


class GetCoord:
    def __init__(self):
        print('加载数据中 ...')
        self.time = []
        self.load_average = []
        self.cpu_percent = []
        self.mem_avail = []
        self.title = ['boot_times', 'top_times', 'time', 'load_average', 'cpus%_us+sy', 'mem_avail', 'app', 'version']

    def read_file(self, log_full_path, csv_full_path):
        L = []
        top_info = []
        build_info = []
        sub_top_info = []  # 存放top信息
        sub_build_info = []  # 存放build信息
        mark_info = 0  # 标志top信息
        mark_info_01 = 0  # 标志build信息
        times_start = 0  # 启动次数
        times_top = 0  # top 次数
        day = 1

        csvfile = open(csv_full_path, 'w', newline='')
        writer = csv.writer(csvfile)
        writer.writerow(self.title)

        with io.open(log_full_path, 'rt', encoding='UTF-8') as fd:  # 打开文件
            for line in fd:  # 按行遍历
                if "Top-Monitor" in line:  # 匹配启动次数
                    if self.time:
                        if sub_top_info:
                            sub_top_info.append(line)
                            top_info.append(copy.deepcopy(sub_top_info))
                            sub_top_info.clear()
                            mark_info = 0
                        times_start = times_start + 1
                        csvfile = open(csv_full_path, 'a+', newline='')
                        writer = csv.writer(csvfile)
                        # print(self.time)
                        alldata = []
                        data = []
                        for index, value in enumerate(self.time):
                            data.append(times_start)
                            data.append(index + 1)
                            data.append(self.time[index])
                            data.append(self.load_average[index])
                            data.append(self.cpu_percent[index])
                            data.append(self.mem_avail[index])
                            # print(top_info)
                            a = ''.join(top_info[index])
                            data.append(a)
                            if index == 0:
                                b = ''.join(build_info[index])
                                data.append(b)

                            alldata.append(copy.deepcopy(data))
                            # print(data)
                            data.clear()

                        for i in alldata:
                            # print(i)
                            # print('写进一次数据')
                            writer.writerow(i)
                        alldata.clear()
                        csvfile.close()
                        self.time.clear()
                        self.load_average.clear()
                        self.cpu_percent.clear()
                        self.mem_avail.clear()
                        self.time.clear()
                        top_info.clear()
                        build_info.clear()
                        day = 1


                    continue
                # if times_start == 3:
                if 'top - ' in line:
                    # print(line)

                    times_top = times_top + 1
                    time_get = re.findall(r"\d{2}:\d{2}:\d{2}", line, re.M)  # 匹配这次top的时间
                    time_get = ''.join(time_get)
                    if self.time:
                        # print(self.time[-1])
                        # print(str(self.time[-1])[2:])
                        time_handle = re.findall(r"\-(.*)", str(self.time[-1]))[0]
                        time_a = datetime.strptime(time_handle, '%H:%M:%S')
                        time_b = datetime.strptime(time_get, '%H:%M:%S')
                        m = (time_b - time_a).seconds
                        if not (12 >= m >= 8):
                            day = day + 1

                        time_get = str(day) + '-' + time_get
                        self.time.append(time_get)  # 获取每个 top 的时间
                        load_average = float(re.findall("\\d+\\.\\d+", line)[0])
                        self.load_average.append(load_average)  # 获取每个 top 的 load_average
                        L.append(load_average)

                    else:
                        time_get = str(day) + '-' + time_get
                        self.time.append(time_get)  # 获取每个 top 的时间
                        load_average = float(re.findall("\\d+\\.\\d+", line)[0])
                        self.load_average.append(load_average)  # 获取每个 top 的 load_average
                if '%Cpu(s)' in line:
                    # print(line)
                    cpu_percent_us = float(re.findall(r"\:(.+?) us", line, re.M)[0])
                    cpu_percent_sy = float(re.findall(r"\,(.+?) sy", line, re.M)[0])
                    cpu_percent = round(cpu_percent_us + cpu_percent_sy, 3)
                    self.cpu_percent.append(cpu_percent)

                if 'KiB Swap:' in line:
                    # print(line)
                    mem_free = int(re.findall(r"\.(.+?) avail Mem", line, re.M)[0])
                    # print(mem_free)
                    self.mem_avail.append(mem_free)
                # if times_start == 5:
                #
                #     break

                if 'PID USER      PR  NI    VIRT    RES    SHR S  %CPU %MEM     TIME+ COMMAND' in line:
                    mark_info = 1
                    # print(line)
                    # print('ok')

                if 'IMAGE_NAME=' in line:
                    mark_info_01 = 1

                if not mark_info == 0:
                    sub_top_info.append(line)
                    mark_info = mark_info + 1
                    if mark_info == 12:
                        top_info.append(copy.deepcopy(sub_top_info))
                        sub_top_info.clear()
                        mark_info = 0

                if not mark_info_01 == 0:
                    sub_build_info.append(line)
                    mark_info_01 = mark_info_01 + 1
                    if mark_info_01 == 6:
                        build_info.append(copy.deepcopy(sub_build_info))
                        sub_build_info.clear()
                        mark_info_01 = 0

            if self.time:
                if sub_top_info:
                    sub_top_info.append(line)
                    top_info.append(copy.deepcopy(sub_top_info))
                    sub_top_info.clear()
                    mark_info = 0
                times_start = times_start + 1
                csvfile = open(csv_full_path, 'a+', newline='')
                writer = csv.writer(csvfile)
                # print(self.time)
                alldata = []
                data = []
                for index, value in enumerate(self.time):
                    data.append(times_start)
                    data.append(index + 1)
                    data.append(self.time[index])
                    data.append(self.load_average[index])
                    data.append(self.cpu_percent[index])
                    data.append(self.mem_avail[index])
                    # print(top_info)
                    a = ''.join(top_info[index])
                    data.append(a)
                    if index == 0:
                        b = ''.join(build_info[index])
                        data.append(b)

                    alldata.append(copy.deepcopy(data))
                    # print(data)
                    data.clear()

                for i in alldata:
                    # print(i)
                    writer.writerow(i)
                alldata.clear()
                csvfile.close()
                self.time.clear()
                self.load_average.clear()
                self.cpu_percent.clear()
                self.mem_avail.clear()
                self.time.clear()
                top_info.clear()
                build_info.clear()
        # for value, index in enumerate(top_info):
        #     print(value)
        #     print(index)


if __name__ == '__main__':
    GetCoord().read_file('static/data/top.log')

