import re
import os
import time
import datetime
import numpy as np
import xlwings as xw

# path = r'\\10.6.32.6\01常熟电性能测试\01.电芯测试\2.原始数据存放（勿删）\金嘉成\CT-LD-2309-03\Shortage\202401-C04739\功率\CC\0℃'
# Cycle_list = os.listdir(path)
#
# list = [datetime.datetime.fromtimestamp(os.path.getctime(os.path.join(path, x))) for x in Cycle_list]
# print(list)
# time_min = min(list)
# time_type = time_min.strftime("%Y/%m/%d %H:%M:%S.%f")
# print(time_type)

# list_1 = [datetime.datetime.fromtimestamp(os.path.getmtime(os.path.join(path, x))) for x in Cycle_list]
# time_min_1 = min(list_1)
# time_type_1 = time_min_1.strftime("%Y/%m/%d %H:%M:%S.%f")
# print(time_type_1)

# list_2 = [datetime.datetime.fromtimestamp(os.path.getatime(os.path.join(path, x))) for x in Cycle_list]
# time_min_2 = min(list_2)
# time_type_2 = time_min_2.strftime("%Y/%m/%d %H:%M:%S.%f")
# print(time_type_2)


all_type = ['水浴', 'DCR', '工况', '充电', '充放电', '快充', '压降监测', '放电', '预处理', 'H/L', 'CAP', 'COR', '试功率', '三电极', '监测压降',
            '能量效率', '放空', '慢充', 'HPPC', 'CC-Rate', '预紧力', '析锂', '电压监控', '试电流', 'CC-HL', '恒功率', '气体体积', 'Cycle', '产气',
            '体积', 'RATE', '水浴循环', '温升', 'K值监测', 'EIS', '自放电', '放电温升', '高温高湿存储', 'ACR', 'DC-Rate', '死体积', '调SOC', 'K值',
            '阶梯放电', '压降监控', '滿充', '高温适应', '产热', 'DC-HL', '活化', '功率', '扎针产气', '低温适应', '产气体积', '满放', '监控压降', 'Storage',
            '低温阶梯放电', '半充', 'BSE', '快充温升', 'OCV', '工况循环', '在线内压', '膨胀力增长', 'CC-DCR']
path_type_a = ['CAP', '满充', 'DC-Rate', 'CC-Rate', '温升', '析锂']
path_type_b = ['功率', 'H/L']

import os


# # 遍历文件夹
# path = r'\\10.6.32.6\01常熟电性能测试\01.电芯测试\2.原始数据存放（勿删）\张振渊\103Ah提起批电解液验证\Shortage\202401-C04749\CAP'
# Cycle_list = os.listdir(path)
#
# file_path = [os.path.abspath(os.path.join(root, file_name)) for root, dirs, files in os.walk(path) for
#              file_name in files if os.path.abspath(os.path.join(root, file_name)).split('.')[-1] == 'xlsx' or
#              os.path.abspath(os.path.join(root, file_name)).split('.')[-1] == 'xls']
# xlsx_min_time = {x: datetime.datetime.fromtimestamp(os.path.getmtime(x)) for x in file_path}
# print(xlsx_min_time)
# xlsx_min_time = min(xlsx_min_time, key=lambda k: xlsx_min_time[k])
# time_xlsx_min = datetime.datetime.fromtimestamp(os.path.getmtime(xlsx_min_time))
#
# print(xlsx_min_time)
# print(time_xlsx_min)


Data_path = r'\\10.6.220.180\tvd实验室\01.电性能测试\01.电芯测试\2.原始数据存放（勿删）\25年之前的原始数据存放（勿删）\沈鑫\CT-LD-2410-23\Shortage\202410-CC10084\CrossTest\800Cycle\小电流测试'
print(os.listdir(Data_path))