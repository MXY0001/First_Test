import os
import tkinter.ttk
from tkinter import *
import xlwings as xw
import tkinter.messagebox
import re
import step_function
import numpy as np
import 工况_图表

# 202310-N01407


class Cycle_Suspended_test_course:
    def __init__(self):
        self.app = xw.App(visible=True, add_book=False)
        self.wb_RPT = self.app.books.open(r'工况RPT数据-Summary-----.xlsx')
        self.Path_CS = r'\\10.6.32.6\01常熟电性能测试\01.电芯测试\1.测试申请单'
        self.Path_NJ = r'\\10.2.2.8\NJtcen\电性能测试\测试申请单'
        self.step_1 = ['搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流放电', '搁置', '恒流充电', '搁置']
        self.step_2 = ['搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流充电', '恒流充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流放电', '搁置']
        self.step_3 = ['搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流放电', '搁置']
        self.step_4 = ['搁置', '恒流放电', '搁置', 'CCCV充电', '搁置', '恒流放电', '搁置', 'CCCV充电', '搁置', '恒流放电', '搁置', 'CCCV充电', '搁置', '恒流放电', '搁置', '搁置', 'CCCV充电', '搁置', '恒流放电', '搁置', '恒流放电', '搁置', '恒流充电', '搁置', '恒流放电', '搁置', '恒流充电', '搁置']
        self.step_5 = ['搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流放电', '搁置', '恒流充电', '搁置', '恒流放电', '搁置', '恒流充电', '搁置']
        self.step_6 = ['搁置', '恒流放电', '搁置', 'CCCV充电', '搁置', '恒流放电', '搁置', 'CCCV充电', '搁置', '恒流放电', '搁置', '恒流充电', '搁置']
        self.step_7 = ['搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流充电', '搁置']
        self.step_8 = ['搁置', '恒流放电', '搁置', 'CCCV充电', '搁置', '恒流放电', '搁置', 'CCCV充电', '搁置', '恒流放电', '搁置', 'CCCV充电', '搁置', '恒流放电', '搁置', '搁置', 'CCCV充电', '搁置', '恒流放电', '搁置', '恒流放电', '搁置', '恒流充电', '搁置']
        self.step_9 = ['搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流放电', '搁置', '恒流充电', '搁置']
        self.step_10 = ['搁置', '恒流放电', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流放电', '搁置', '恒流充电', '搁置', '恒流放电', '搁置', '恒流充电', '搁置']
        self.step_11 = ['搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流放电', '搁置', '搁置']
        self.step_12 = ['恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流放电', '搁置', '恒流充电', '搁置']
        self.step_13 = ['搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流放电', '搁置', '恒流充电', '搁置']
        self.step_14 = ['搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置']
        self.step_15 = ['搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流放电', '搁置', '恒流充电', '搁置']
        self.step_16 = ['搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流放电', '搁置']
        self.step_17 = ['搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流放电']
        self.step_18 = ['搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流放电', '搁置']
        self.step_19 = ['搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流放电', '搁置']
        self.step_20 = ['搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流恒压充电', '搁置', '恒流放电', '搁置', '恒流放电', '搁置']

        self.step_sum = [self.step_1, self.step_2, self.step_3, self.step_4, self.step_5, self.step_6, self.step_7, self.step_8, self.step_9, self.step_10,
                         self.step_11, self.step_12, self.step_13, self.step_14, self.step_15, self.step_16, self.step_17, self.step_18, self.step_19, self.step_20]


        self.root = Tk()
        self.root.title('小程序')
        self.root.geometry('200x130')
        self.root.attributes('-topmost', True)
        self.progressbarOne = tkinter.ttk.Progressbar(self.root)
        self.progressbarOne.place(x=20, y=60, width=150)
        self.btn1 = Button(self.root, text="运行", command=self.Main)
        self.btn1.place(x=100, y=90)
        self.btn2 = Button(self.root, text="清除", command=self.clear)
        self.btn2.place(x=20, y=90)
        self.btn3 = Button(self.root, text="条码", command=self.Barcode_write)
        self.btn3.place(x=60, y=90)
        self.btn4 = Button(self.root, text="图表", command=lambda: 工况_图表.Chart().main())
        self.btn4.place(x=140, y=90)
        self.label_0 = Label(self.root, text='请输入测试单号：')
        self.label_0.place(x=20, y=0)
        self.entry0 = Entry(self.root)
        self.entry0.place(x=20, y=30, width=150)
        self.root.mainloop()

    # 把字符类型时间转化为浮点数 单位时间 员工工衣尺码
    @staticmethod
    def Time_To_Float(Time):

        Time = str(Time)
        Days = int(Time.split('.')[0])
        Hours = int(Time.split(':')[0][-2:])
        Minutes = int(Time.split(':')[1])
        Seconds = float(Time.split(':')[2])
        Time_Hours = Days*24+Hours+(Seconds/60+Minutes)/60

        return Time_Hours

    @staticmethod
    def Get_slope_charge(V_value_record_1, charge_cap_value):
        a = [float(x) for x in list(V_value_record_1)]
        b = [float(x) for x in list(charge_cap_value)]
        a = [a[x:x+20] for x in range(len(a))]
        b = [b[x:x+20] for x in range(len(a))]

        ab = [sum([a[x][i] * b[x][i] for i in range(len(a[x]))]) * len(a[x]) for x in range(len(a))]
        a_sum_b_sum = [sum(a[x]) * sum(b[x]) for x in range(len(a))]
        bb_sum = [sum([b[x][i] ** 2 for i in range(len(a[x]))]) * len(a[x]) for x in range(len(a))]
        b_sum_b_sum = [sum(b[x]) ** 2 for x in range(len(a))]

        aa_sum = [sum([a[x][i] ** 2 for i in range(len(a[x]))]) * len(a[x]) for x in range(len(a))]
        a_sum_a_sum = [sum(a[x]) ** 2 for x in range(len(a))]

        result_1 = [(ab[x] - a_sum_b_sum[x]) / (bb_sum[x] - b_sum_b_sum[x]) if (bb_sum[x]-b_sum_b_sum[x]) != 0 else 0 for x in range(len(a))]
        result_2 = [(ab[x] - a_sum_b_sum[x]) / (aa_sum[x] - a_sum_a_sum[x]) if (aa_sum[x]-a_sum_a_sum[x]) != 0 else 0 for x in range(len(a))]
        return result_1, result_2

    @staticmethod
    def Get_slope_discharge(V_value_record_2, charge_cap_value_1):
        a = [float(x) for x in list(V_value_record_2)]
        b = [float(x) for x in list(charge_cap_value_1)]
        a = [a[x:x+20] for x in range(len(a))]
        b = [b[x:x+20] for x in range(len(a))]

        ab = [sum([a[x][i] * b[x][i] for i in range(len(a[x]))]) * len(a[x]) for x in range(len(a))]
        a_sum_b_sum = [sum(a[x]) * sum(b[x]) for x in range(len(a))]
        bb_sum = [sum([b[x][i] ** 2 for i in range(len(a[x]))]) * len(a[x]) for x in range(len(a))]
        b_sum_b_sum = [sum(b[x]) ** 2 for x in range(len(a))]

        aa_sum = [sum([a[x][i] ** 2 for i in range(len(a[x]))]) * len(a[x]) for x in range(len(a))]
        a_sum_a_sum = [sum(a[x]) ** 2 for x in range(len(a))]

        result_1 = [(ab[x] - a_sum_b_sum[x]) / (bb_sum[x] - b_sum_b_sum[x]) if (bb_sum[x]-b_sum_b_sum[x]) != 0 else 0 for x in range(len(a))]
        result_1 = [abs(x) for x in result_1]
        result_2 = [(ab[x] - a_sum_b_sum[x]) / (aa_sum[x] - a_sum_a_sum[x]) if (aa_sum[x]-a_sum_a_sum[x]) != 0 else 0 for x in range(len(a))]
        result_2 = [abs(x) for x in result_2]
        return result_1, result_2

    # 获取对应测试单超链接
    def Get_hyperlink(self, app_data, sht_RPT_0):
        Csd_No = sht_RPT_0.range('c1').value
        Year = Csd_No[0:4] + '年'
        Month_1 = Csd_No[0:6]
        Month_2 = Csd_No[4:6].lstrip('0') + '月'
        Excel_path = 0
        if Csd_No.split('-')[1][0] == 'C':
            Excel_path = self.Path_CS + '\\' + Year + '\\' + Month_1
        elif Csd_No.split('-')[1][0] == 'N' or Csd_No.split('-')[1][0] == '0':
            if Csd_No[0:4] == '2023' or Csd_No[0:4] == '2024':
                Excel_path = self.Path_NJ + '\\' + Year + '\\' + Month_2
            elif Csd_No[0:4] == '2021' or Csd_No[0:4] == '2022':
                Excel_path = self.Path_NJ + '\\' + Year + '\\' + Month_1
        elif Csd_No.split('-')[1][0] == 'O':
            self.Path_CS = r'Z:\常熟研发测试数据\01.电芯测试\10.ORT测试申请单\ORT测试申请单'
            Excel_path = self.Path_CS + '\\' + Year

        if Excel_path == 0:
            return '0'
        excel_list = os.listdir(Excel_path)
        Excel_name = [x for x in excel_list if Csd_No in x and x[0:2] != '~$']
        if len(Excel_name) == 0:
            return '0'
        Excel_name_path = os.path.join(Excel_path, Excel_name[0])

        wb_Csd = app_data.books.open(Excel_name_path)
        sht_names = wb_Csd.sheet_names

        try:
            sht_csbg = wb_Csd.sheets['测试报告']
        except Exception:
            sht_csbg = 0
            for x in sht_names:
                if '测试报告' in x:
                    sht_csbg = wb_Csd.sheets[x]
                    break
            if sht_csbg == 0:
                wb_Csd.close()
                return '1'

        B_value = sht_csbg[0:20, 1:2].value
        K_value = sht_csbg[0:20, 10:11].value
        index_hyperlink = 7

        # if '新Data/raw' in K_value:
        #     index_hyperlink = K_value.index('新Data/raw')
        # elif '工况' in B_value:
        #     index_hyperlink = B_value.index('工况')
        # elif '循环测试' in B_value:
        #     index_hyperlink = B_value.index('循环测试')
        # elif'Cyclelife' in B_value:
        #     index_hyperlink = B_value.index('Cyclelife')
        # else:
        for x in range(20):
            if B_value[x] is not None and '工况' in B_value[x]:
                index_hyperlink = x
                break

        try:
            hyperlink = sht_csbg.range('k' + str(index_hyperlink + 1)).hyperlink
            if '../../../' in hyperlink:
                hyperlink = hyperlink.replace(r'../../../', '\\\\10.6.32.6\\01常熟电性能测试\\01.电芯测试\\')
            path = os.listdir(hyperlink)
        except Exception:
            wb_Csd.close()
            return '2'


        wb_Csd.close()


        return hyperlink

    def Get_barcode(self, sht_RPT_0):
        Csd_No = sht_RPT_0.range('c1').value
        Year = Csd_No[0:4] + '年'
        Month_1 = Csd_No[0:6]
        Month_2 = Csd_No[4:6].lstrip('0') + '月'
        Excel_path = 0
        if Csd_No.split('-')[1][0] == 'C':
            Excel_path = self.Path_CS + '\\' + Year + '\\' + Month_1
        elif Csd_No.split('-')[1][0] == 'N' or Csd_No.split('-')[1][0] == '0':
            if Csd_No[0:4] == '2023' or Csd_No[0:4] == '2024':
                Excel_path = self.Path_NJ + '\\' + Year + '\\' + Month_2
            elif Csd_No[0:4] == '2021' or Csd_No[0:4] == '2022':
                Excel_path = self.Path_NJ + '\\' + Year + '\\' + Month_1
        if Excel_path == 0:
            return 0, 0
        excel_list = os.listdir(Excel_path)
        Excel_name = [x for x in excel_list if Csd_No in x and x[0:2] != '~$']
        if len(Excel_name) == 0:
            return 0, 1
        Excel_name_path = os.path.join(Excel_path, Excel_name[0])

        app_data = xw.App(visible=False, add_book=False)
        wb_Csd = app_data.books.open(Excel_name_path)
        sht_names = wb_Csd.sheet_names
        try:
            sht_csbg = wb_Csd.sheets['测试报告']
        except Exception:
            sht_csbg = 0
            for x in sht_names:
                if '测试报告' in x:
                    sht_csbg = wb_Csd.sheets[x]
                    break
            if sht_csbg == 0:
                wb_Csd.close()
                return 0, 2

        B_value = sht_csbg[0:20, 1:2].value
        B_value = ['ww' if x is None else x for x in B_value]
        K_value = sht_csbg[0:20, 10:11].value
        index_hyperlink = 7
        if '新Data/raw' in K_value:
            index_hyperlink = K_value.index('新Data/raw')
        elif '工况' in B_value:
            index_hyperlink = B_value.index('工况')
        else:
            for x in range(20):
                if '工况' in B_value[x]:
                    index_hyperlink = x
                    break
        try:
            hyperlink = sht_csbg.range('k' + str(index_hyperlink + 1)).hyperlink
            path = os.listdir(hyperlink)
        except Exception:
            tkinter.messagebox.showinfo(message= '超链接地址错误')
            wb_Csd.close()
            return 0, 3

        if 'Summary' in sht_names:
            sht_Summary = wb_Csd.sheets['Summary']
            row_max = sht_Summary.used_range.last_cell.row
            A_value = sht_Summary[0:row_max, 0:1].value
            B_value = sht_Summary[0:row_max, 1:2].value
            if 'NO.' in A_value:
                index_No = A_value.index('NO.')
                row_value = sht_Summary[index_No:index_No + 1, 0:1].expand('table').rows.count-1
                Barcode_value = sht_Summary[index_No+1:index_No+1+row_value, 0:14].value
            elif 'NO.' in B_value:
                index_No = B_value.index('NO.')
                row_value = sht_Summary[index_No:index_No + 1, 1:2].expand('table').rows.count-1
                Barcode_value = sht_Summary[index_No + 1:index_No + 1 + row_value, 1:15].value
            else:
                index_No = 3
                row_value = sht_Summary[index_No:index_No + 1, 0:1].expand('table').rows.count-1
                Barcode_value = sht_Summary[index_No + 1:index_No + 1 + row_value, 0:14].value

            wb_Csd.close()
            app_data.quit()
            return Barcode_value, hyperlink
        else:
            sht_Summary = 0
            for x in sht_names:
                if 'Summary' in x:
                    sht_Summary = wb_Csd.sheets[x]
                    row_max = sht_Summary.used_range.last_cell.row
                    A_value = sht_Summary[0:row_max, 0:1].value
                    B_value = sht_Summary[0:row_max, 1:2].value
                    if 'NO.' in A_value:
                        index_No = A_value.index('NO.')
                        row_value = sht_Summary[index_No:index_No + 1, 0:1].expand('table').rows.count - 1
                        Barcode_value = sht_Summary[index_No + 1:index_No + 1 + row_value, 0:14].value
                    elif 'NO.' in B_value:
                        index_No = B_value.index('NO.')
                        row_value = sht_Summary[index_No:index_No + 1, 1:2].expand('table').rows.count - 1
                        Barcode_value = sht_Summary[index_No + 1:index_No + 1 + row_value, 1:15].value
                    else:
                        index_No = 3
                        row_value = sht_Summary[index_No:index_No + 1, 0:1].expand('table').rows.count - 1
                        Barcode_value = sht_Summary[index_No + 1:index_No + 1 + row_value, 0:14].value

                    wb_Csd.close()
                    app_data.quit()
                    return Barcode_value, hyperlink
            if sht_Summary == 0:
                wb_Csd.close()
                app_data.quit()
                return 0, 4

    # 获取表工步信息中的数据
    def Get_step_message(self, wb_data):
        try:
            sht_step = wb_data.sheets['工步信息']
        except Exception:
            sht_step = wb_data.sheets[2]
        col_step = sht_step.used_range.last_cell.column
        row_recode = sht_step.range('a1').expand('table').rows.count - 1
        title_record = sht_step[0:1, 0:col_step].value
        try:
            index_step = title_record.index('工步名称')
        except Exception:
            index_step = title_record.index('状态')
        step_value = sht_step[1:1+row_recode, index_step:index_step+1].value
        try:
            step_index = self.step_sum.index(step_value)
        except ValueError:
            # tkinter.messagebox.showinfo(message='未录入该工步')
            return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
        if step_index == 0:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_1(wb_data, 0)
        elif step_index == 1:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_2(wb_data, 1)
        elif step_index == 2:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_2(wb_data, 2)
        elif step_index == 3:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_1(wb_data, 3)
        elif step_index == 4:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_1(wb_data, 4)
        elif step_index == 5:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_3(wb_data, 5)
        elif step_index == 6:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_3(wb_data, 6)
        elif step_index == 7:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_1(wb_data, 7)
        elif step_index == 8:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_1(wb_data, 8)
        elif step_index == 9:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_4(wb_data, 9)
        elif step_index == 10:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_1(wb_data, 10)
        elif step_index == 11:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_4(wb_data, 11)
        elif step_index == 12:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_2(wb_data, 12)
        elif step_index == 13:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_2(wb_data, 13)
        elif step_index == 14:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_1(wb_data, 14)
        elif step_index == 15:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_5(wb_data, 15)
        elif step_index == 16:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_5(wb_data, 16)
        elif step_index == 17:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_6(wb_data, 17)
        elif step_index == 18:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_2(wb_data, 18)
        elif step_index == 19:
            charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
            CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
            charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = step_function.step_7(wb_data, 19)

        else:
            return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

        return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
                CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
                charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1

    # 获取表 记录数据 中的数据
    @staticmethod
    def Get_record_message(wb_data):
        try:
            sht_record = wb_data.sheets['记录数据']
        except  Exception:
            try:
                sht_record = wb_data.sheets[2]
            except Exception:
                return 0
        col_record = sht_record.used_range.last_cell.column
        row_record = sht_record.range('a1').expand('table').rows.count
        title_record = sht_record[0:1, 0:col_record].value
        try:
            index_Absolute_time = title_record.index('绝对时间')
        except ValueError:
            for x in range(len(title_record)):
                if '绝对时间' in title_record[x]:
                    index_Absolute_time = x
                    break

        Test_time_start = sht_record[1:2, index_Absolute_time:index_Absolute_time+1].value
        Test_time_end = sht_record[row_record-1:row_record, index_Absolute_time:index_Absolute_time+1].value

        return Test_time_start, Test_time_end


    def Main(self):
        app_data = xw.App(visible=True, add_book=False)
        self.wb_RPT = xw.books.active
        sht_RPT_0 = self.wb_RPT.sheets['RPT数据']
        sht_Summary = self.wb_RPT.sheets['Summary']
        row_Summary = sht_Summary.range('b6').expand('table').rows.count - 1
        if row_Summary == 0:
            tkinter.messagebox.showinfo(message='请点击条码，将barcode填入Summary')
            return
        elif row_Summary == 1:
            barcode_value = [sht_Summary[6:6+row_Summary, 3:4].value]
        else:
            barcode_value = sht_Summary[6:6 + row_Summary, 3:4].value

        row_RPT = sht_RPT_0.used_range.last_cell.row
        C_value = sht_RPT_0[0:row_RPT, 2:3].value
        hyperlink = self.Get_hyperlink(app_data, sht_RPT_0)
        if hyperlink == '0':
            tkinter.messagebox.showinfo(message='未找到测试单所在路径')
            return
        elif hyperlink == '1':
            tkinter.messagebox.showinfo(message='测试单无‘测试报告或Summary’sheet页')
            return
        elif hyperlink == '2':
            tkinter.messagebox.showinfo(message='超链接打不开')
            return
        hyperlink_list = os.listdir(hyperlink)
        if 'Cross Test' in hyperlink_list:
            Cycle_path = hyperlink+'\\'+'Cross Test'
            Cycle_list = os.listdir(Cycle_path)
        elif 'CrossTest' in hyperlink_list:
            Cycle_path = hyperlink+'\\'+'CrossTest'
            Cycle_list = os.listdir(Cycle_path)
        elif 'CorssTest' in hyperlink_list:
            Cycle_path = hyperlink + '\\' + 'CorssTest'
            Cycle_list = os.listdir(Cycle_path)
        elif 'Corss Test' in hyperlink_list:
            Cycle_path = hyperlink+'\\'+'Corss Test'
            Cycle_list = os.listdir(Cycle_path)
        else:
            Cycle_path = hyperlink
            Cycle_list = os.listdir(Cycle_path)

        Cycle_list = ['0cls' if 'BF' in x else str(x) for x in Cycle_list]
        Cycle_list = ['99999999cls' if 'AF' in str(x) else x for x in Cycle_list]
        Cycle_list = sorted(Cycle_list, key=lambda file: int(re.search(r"(-?\d+)", file).group()))
        Cycle_list = ['BF' if x.strip() == '0cls' else x for x in Cycle_list]
        Cycle_list = ['AF' if x.strip() == '99999999cls' else x for x in Cycle_list]
        # Cycle_list = [int(re.search(r"(-?\d+)", x).group()) if x != '99999999cls' else x for x in Cycle_list]

        self.progressbarOne['maximum'] = len(Cycle_list)
        for i in range(len(Cycle_list)):
            sht_RPT = self.wb_RPT.sheets['RPT数据']
            row_summary_0 = sht_Summary.used_range.last_cell.row
            D_Summary = sht_Summary[0:row_summary_0, 3:4].value
            if Cycle_list[i].strip() == 'BF':
                Cycle_No = 0
                Cycle_No_j = 0
            elif str(Cycle_list[i]).strip() == 'AF':
                Cycle_No = 'AF'
                Cycle_No_j = 9000000000
            else:
                # Cycle_No = re.search(r"(-?\d+)", Cycle_list[i]).group()
                Cycle_No = Cycle_list[i]
                Cycle_No_j = re.search(r"(-?\d+)", Cycle_list[i]).group()


            Data_path = os.path.join(Cycle_path, Cycle_list[i])
            Data_list = os.listdir(Data_path)


            if 'DCR' in Data_list:
                Data_path = os.path.join(Data_path, 'DCR')
                Data_list = os.listdir(Data_path)
            elif 'CAP' in Data_list:
                    Data_path = os.path.join(Data_path, 'CAP')
                    Data_list = os.listdir(Data_path)

            Data_list = [x for x in Data_list if x[0:2] != '~$' and x[-4:] == 'xlsx']
            try:
                Data_Barcode = [x.split('-')[3] for x in Data_list]
            except IndexError:
                Data_Barcode = []
            Data_Barcode_1 = [x.split('-')[0] for x in Data_list]
            try:
                Data_Barcode_2 = ['-'.join(x.split('-')[3:6]) for x in Data_list]
            except IndexError:
                Data_Barcode_2 = []
            for j in range(len(Data_list)):
                if Data_Barcode != [] and Data_Barcode[j] in barcode_value:
                    barcode_names = 'DVDQ-' + str(Data_Barcode[j])
                elif Data_Barcode_1[j] in barcode_value:
                    barcode_names = 'DVDQ-' + str(Data_Barcode_1[j])
                elif Data_Barcode_2 != [] and Data_Barcode_2[j] in barcode_value:
                    barcode_names = 'DVDQ-' + str(Data_Barcode_2[j])
                else:
                    continue
                try:
                    sht_Cycle_list = self.wb_RPT.sheets[barcode_names]
                except Exception:
                    sheet_namaes = self.wb_RPT.sheet_names
                    for x in sheet_namaes:
                        if barcode_names in x:
                            barcode_names = x
                            break

                    sht_Cycle_list = self.wb_RPT.sheets[barcode_names]

                col_Cycle_list = sht_Cycle_list.range('b2').expand('table').columns.count
                # row_Cycle_list = sht_Cycle_list.used_range.last_cell.row
                barcode_Cycle_list = sht_Cycle_list[0:1, 0:col_Cycle_list].value
                try:
                    barcode_index = C_value.index(Data_Barcode[j])
                except ValueError:
                    try:
                        barcode_index = C_value.index(Data_Barcode_1[j])
                    except ValueError:
                        continue
                except IndexError:
                    try:
                        barcode_index = C_value.index(Data_Barcode_1[j])
                    except ValueError:
                        continue
                try:
                    barcode_index_1 = barcode_Cycle_list.index(Cycle_list[i])
                except ValueError:
                    try:
                        col_end = sht_Cycle_list.range('a2').expand('table').columns.count
                        sht_Cycle_list[0:1, col_end:col_end+1].value = Cycle_list[i]
                        sht_Cycle_list[1:2, col_end:col_end+1].value = '充电容量(Ah)'
                        sht_Cycle_list[1:2, col_end+1:col_end+2].value = '电压（V)'
                        sht_Cycle_list[1:2, col_end+2:col_end+3].value = 'DV/DQ'
                        sht_Cycle_list[1:2, col_end+3:col_end+4].value = 'DQ/DV'
                        sht_Cycle_list[1:2, col_end+4:col_end+5].value = '放电容量(Ah)'
                        sht_Cycle_list[1:2, col_end+5:col_end+6].value = '电压（V)'
                        sht_Cycle_list[1:2, col_end+6:col_end+7].value = 'DV/DQ'
                        sht_Cycle_list[1:2, col_end+7:col_end+8].value = 'DQ/DV'
                        # 设置格式
                        sht_Cycle_list[0:1, col_end:col_end+8].api.Merge()
                        sht_Cycle_list[0:1, col_end:col_end+8].api.HorizontalAlignment = -4108
                        sht_Cycle_list[0:1, col_end:col_end+8].api.VerticalAlignment = -4108
                        sht_Cycle_list[0:, col_end:col_end+8].api.Borders(7).LineStyle = 1
                        sht_Cycle_list[0:, col_end:col_end+8].api.Borders(7).Weight = 3
                        sht_Cycle_list[1:2, col_end:col_end+2].api.Font.Bold = True
                        sht_Cycle_list[1:2, col_end:col_end+2].column_width = 13.5
                        sht_Cycle_list[1:2, col_end:col_end+4].color = (75, 172, 198)
                        sht_Cycle_list[1:2, col_end+4:col_end+8].api.Font.Bold = True
                        sht_Cycle_list[1:2, col_end+4:col_end+8].column_width = 13.5
                        sht_Cycle_list[1:2, col_end+4:col_end+8].color = (155, 187, 89)
                        col_Cycle_list = sht_Cycle_list.range('b2').expand('table').columns.count
                        barcode_Cycle_list = sht_Cycle_list[0:1, 0:col_Cycle_list].value
                        barcode_index_1 = barcode_Cycle_list.index(Cycle_list[i])
                    except ValueError:
                        continue

                # row_Cycle_No = i
                row_Cycle_No = sht_RPT[barcode_index+2:barcode_index+3, 2:3].expand('table').rows.count - 1
                if sht_RPT[barcode_index + 3+i:barcode_index + 4+i, 2:3].value is not None: # 若已经对应值跳过
                    continue
                if row_Cycle_No == 1 or row_Cycle_No == 0:
                    list_cycle = [sht_RPT[barcode_index + 3:barcode_index + 3 + row_Cycle_No, 2:3].value]
                else:
                    list_cycle = sht_RPT[barcode_index + 3:barcode_index + 3 + row_Cycle_No, 2:3].value
                    list_cycle = [int(x) if type(x) != str else x for x in list_cycle]

                if Cycle_No_j == 'AF':
                    Cycle_No_j = 9000000000
                if int(Cycle_No_j) in list_cycle: # or int(Cycle_No_j) == 9000000000
                    continue

                try:
                    wb_data = app_data.books.open(os.path.join(Data_path, Data_list[j]))
                except Exception:
                    tkinter.messagebox.showinfo(message=os.path.join(Data_path, Data_list[j])+'  打不开')
                    continue
                Test_time = self.Get_record_message(wb_data)
                if Test_time == 0:
                    wb_data.close()
                    continue
                charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, discharge_standing_end_V, \
                CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, discharge_end_I_10, standing_end2_V, charge_end_V_10, charge_end_I_10, \
                charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress, V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1 = self.Get_step_message(wb_data)

                if charge_cap == 0 or sht_RPT[barcode_index + 2 + row_Cycle_No:barcode_index + 2 + row_Cycle_No, 2:3].value == Cycle_No:
                    wb_data.close()
                    continue

                if i == 0:
                    if Data_Barcode != [] and Data_Barcode[j] in D_Summary:
                        barcode_index_0 = D_Summary.index(Data_Barcode[j])
                        if sht_Summary[barcode_index_0:barcode_index_0+1, 8:9].value is None and sht_Summary[barcode_index_0:barcode_index_0+1, 8:9].api.Font.Color != 255.0:
                            sht_Summary[barcode_index_0:barcode_index_0 + 1, 8:9].value = discharge_cap
                            sht_Summary[barcode_index_0:barcode_index_0 + 1, 11:12].value = Test_time[1]
                    elif Data_Barcode_1 != [] and Data_Barcode_1[j] in D_Summary:
                        barcode_index_0 = D_Summary.index(Data_Barcode_1[j])
                        if sht_Summary[barcode_index_0:barcode_index_0+1, 8:9].value is None and sht_Summary[barcode_index_0:barcode_index_0+1, 8:9].api.Font.Color != 255.0:
                            sht_Summary[barcode_index_0:barcode_index_0 + 1, 8:9].value = discharge_cap
                            sht_Summary[barcode_index_0:barcode_index_0 + 1, 11:12].value = Test_time[1]

                if Data_Barcode != [] and Data_Barcode[j] in D_Summary:
                    barcode_index_0 = D_Summary.index(Data_Barcode[j])
                    Cycle_Summary = sht_Summary[barcode_index_0:barcode_index_0+1, 7:8].value
                    if Cycle_Summary is None or Cycle_Summary <= int(Cycle_No_j):
                        if Cycle_No_j == 9000000000:
                            Cycle_No_j = 'AF'
                        sht_Summary[barcode_index_0:barcode_index_0 + 1, 7:8].value = Cycle_No_j
                        sht_Summary[barcode_index_0:barcode_index_0 + 1, 9:10].value = discharge_cap
                        sht_Summary[barcode_index_0:barcode_index_0 + 1, 12:13].value = Test_time[1]
                elif Data_Barcode_1[j] in D_Summary:
                    barcode_index_0 = D_Summary.index(Data_Barcode_1[j])
                    Cycle_Summary = sht_Summary[barcode_index_0:barcode_index_0+1, 7:8].value
                    if Cycle_Summary is None or Cycle_Summary <= int(Cycle_No_j):
                        if Cycle_No_j == 9000000000:
                            Cycle_No_j = 'AF'
                        sht_Summary[barcode_index_0:barcode_index_0 + 1, 7:8].value = Cycle_No_j
                        sht_Summary[barcode_index_0:barcode_index_0 + 1, 9:10].value = discharge_cap
                        sht_Summary[barcode_index_0:barcode_index_0 + 1, 12:13].value = Test_time[1]
                if Cycle_No_j == 9000000000:
                    Cycle_No_j = 'AF'
                sht_RPT[barcode_index + 3 + row_Cycle_No:barcode_index + 4 + row_Cycle_No, 1:2].value = [Test_time[0], Cycle_No_j, charge_cap, charge_energy, discharge_cap, discharge_energy]
                sht_RPT[barcode_index + 3 + row_Cycle_No:barcode_index + 4 + row_Cycle_No, 12:13].value = self.Time_To_Float(CC_time)
                sht_RPT[barcode_index + 3 + row_Cycle_No:barcode_index + 4 + row_Cycle_No, 14:15].value = [self.Time_To_Float(total_time), CC_cap]
                sht_RPT[barcode_index + 3 + row_Cycle_No:barcode_index + 4 + row_Cycle_No, 21:22].value = [charge_end_V, charge_standing_end_V]
                sht_RPT[barcode_index + 3 + row_Cycle_No:barcode_index + 4 + row_Cycle_No, 24:25].value = [discharge_end_V, discharge_standing_end_V]
                sht_RPT[barcode_index + 3 + row_Cycle_No:barcode_index + 4 + row_Cycle_No, 27:28].value = [standing_end1_V, discharge_end_V_10, discharge_end_I_10]
                sht_RPT[barcode_index + 3 + row_Cycle_No:barcode_index + 4 + row_Cycle_No, 31:32].value = [standing_end2_V, charge_end_V_10, charge_end_I_10]
                sht_RPT[barcode_index + 3 + row_Cycle_No:barcode_index + 4 + row_Cycle_No, 35:36].value = [charge_max_stress, discharge_max_stress]
                sht_RPT[barcode_index + 3 + row_Cycle_No:barcode_index + 4 + row_Cycle_No, 37:38].value = [charge_max_tem, discharge_max_tem]
                charge_cap_value_1 = list(charge_cap_value_1)
                charge_cap_value_1 = [abs(float(x)) for x in charge_cap_value_1]

                sht_Cycle_list[2:3, barcode_index_1:barcode_index_1+1].options(transpose=True).value = charge_cap_value
                sht_Cycle_list[2:3, barcode_index_1+1:barcode_index_1+2].options(transpose=True).value = V_value_record_1
                sht_Cycle_list[2:3, barcode_index_1+4:barcode_index_1+5].options(transpose=True).value = charge_cap_value_1
                sht_Cycle_list[2:3, barcode_index_1+5:barcode_index_1+6].options(transpose=True).value = V_value_record_2
                sht_Cycle_list[2:3, barcode_index_1+2:barcode_index_1+3].options(transpose=True).value = self.Get_slope_charge(V_value_record_1, charge_cap_value)[0]
                sht_Cycle_list[2:3, barcode_index_1+3:barcode_index_1+4].options(transpose=True).value = self.Get_slope_charge(V_value_record_1, charge_cap_value)[1]
                sht_Cycle_list[2:3, barcode_index_1+6:barcode_index_1+7].options(transpose=True).value = self.Get_slope_discharge(V_value_record_2, charge_cap_value_1)[0]
                sht_Cycle_list[2:3, barcode_index_1+7:barcode_index_1+8].options(transpose=True).value = self.Get_slope_discharge(V_value_record_2, charge_cap_value_1)[1]
                sht_Cycle_list[2:2+len(charge_cap_value), barcode_index_1:barcode_index_1+2].api.NumberFormat = '0.0000'
                sht_Cycle_list[2:2+len(charge_cap_value_1), barcode_index_1+4:barcode_index_1+6].api.NumberFormat = '0.0000'
                sht_Cycle_list[1:1+len(charge_cap_value), barcode_index_1:barcode_index_1+2].api.HorizontalAlignment = -4108
                sht_Cycle_list[1:1+len(charge_cap_value), barcode_index_1:barcode_index_1+2].api.VerticalAlignment = -4108
                sht_Cycle_list[1:1+len(charge_cap_value_1), barcode_index_1+4:barcode_index_1+6].api.HorizontalAlignment = -4108
                sht_Cycle_list[1:1+len(charge_cap_value_1), barcode_index_1+4:barcode_index_1+6].api.VerticalAlignment = -4108
                wb_data.close()

            self.progressbarOne['value'] = i + 1
            self.root.update()
            if i == len(Cycle_list) - 1:
                app_data.quit()
                tkinter.messagebox.showinfo(message='完成')

    def clear(self):
        工况_图表.Chart().clear()
        self.wb_RPT = xw.books.active
        sheet_names = self.wb_RPT.sheet_names
        sht_Summary = self.wb_RPT.sheets['Summary']
        row_Summary = sht_Summary.range('b6').expand('table').rows.count - 1
        if row_Summary == 0:
            value_Summary = []
            tkinter.messagebox.showinfo(message='Summary Barcode 为空')
        elif row_Summary == 1:
            value_Summary = [sht_Summary[6:6+row_Summary, 3:4].value]
        else:
            value_Summary = sht_Summary[6:6 + row_Summary, 3:4].value
        for x in sheet_names:
            if 'RPT数据' in x or 'Summary' in x:
                continue
            else:
                row_1 = self.wb_RPT.sheets[x].used_range.last_cell.row
                col_1 = self.wb_RPT.sheets[x].used_range.last_cell.row
                try:
                    self.wb_RPT.sheets[x][2:row_1, 1:col_1].clear_contents()
                    self.wb_RPT.sheets[x][0:1, 1:col_1].clear_contents()
                except Exception:
                    pass
                if x.split('-')[1].strip() in value_Summary:
                    continue
                else:
                    self.wb_RPT.sheets[x].delete()
                # self.wb_RPT.sheets[x].delete()
        sht_RPT = self.wb_RPT.sheets['RPT数据']
        row_RPT = sht_RPT.used_range.last_cell.row
        B_value = sht_RPT[0:row_RPT, 1:2].value
        B_value_np = np.array(B_value)
        recode = np.where(B_value_np == '测试时间')[0]
        index_list = list(recode)
        sht_RPT[0:1, 2:3].clear_contents()
        sht_RPT[547:, 0:40].clear_contents()


        for i in range(len(index_list)):
            sht_RPT[index_list[i] - 2:index_list[i] - 1, 2:3].clear_contents()
            sht_RPT[index_list[i] + 1:index_list[i]+31, 1:7].clear_contents()
            sht_RPT[index_list[i] + 1:index_list[i]+31, 12:13].clear_contents()
            sht_RPT[index_list[i] + 1:index_list[i]+31, 14:16].clear_contents()
            sht_RPT[index_list[i] + 1:index_list[i]+31, 21:23].clear_contents()
            sht_RPT[index_list[i] + 1:index_list[i]+31, 24:26].clear_contents()
            sht_RPT[index_list[i] + 1:index_list[i]+31, 27:30].clear_contents()
            sht_RPT[index_list[i] + 1:index_list[i]+31, 31:34].clear_contents()
            sht_RPT[index_list[i] + 1:index_list[i]+31, 35:39].clear_contents()


    def Barcode_write(self):
        self.wb_RPT = xw.books.active
        app_data = xw.App(visible=True, add_book=False)
        sheet_names = self.wb_RPT.sheet_names
        sht_RPT = self.wb_RPT.sheets['RPT数据']
        sht_Summary = self.wb_RPT.sheets['Summary']
        sht_RPT[0:1, 2:3].value = self.entry0.get()
        hyperlink = self.Get_hyperlink(app_data, self.wb_RPT.sheets['RPT数据'])
        row = sht_Summary.range('b6').expand('table').rows.count - 1
        if row == 1:
            Barcode_value = [sht_Summary[6:6+row, 3:4].value]
        else:
            Barcode_value = sht_Summary[6:6+row, 3:4].value

        for i in range(len(Barcode_value)):
            sht_RPT[1+34*i:2+34*i, 2:3].value = Barcode_value[i]

        hyperlink_list = os.listdir(hyperlink)
        if 'Cross Test' in hyperlink_list:
            Cycle_path = hyperlink + '\\' + 'Cross Test'
            Cycle_list = os.listdir(Cycle_path)
        elif 'CrossTest' in hyperlink_list:
            Cycle_path = hyperlink + '\\' + 'CrossTest'
            Cycle_list = os.listdir(Cycle_path)
        elif 'Corss Test' in hyperlink_list:
            Cycle_path = hyperlink + '\\' + 'Corss Test'
            Cycle_list = os.listdir(Cycle_path)
        elif  'CorssTest' in hyperlink_list:
            Cycle_path = hyperlink + '\\' + 'CorssTest'
            Cycle_list = os.listdir(Cycle_path)
        else:
            tkinter.messagebox.showinfo(message='Cross Test 不存在或命名不规范')
            return
            # Cycle_path = hyperlink
            # Cycle_list = os.listdir(Cycle_path)
        Cycle_list = ['0cls' if 'BF' in str(x) else x for x in Cycle_list]
        Cycle_list = ['99999999cls' if 'AF' in str(x) else x for x in Cycle_list]
        Cycle_list = sorted(Cycle_list, key=lambda file: int(re.search(r"(-?\d+)", file).group()))
        Cycle_list = ['BF' if x.strip() == '0cls' else x for x in Cycle_list]
        Cycle_list = ['AF' if x.strip() == '99999999cls' else x for x in Cycle_list]
        self.progressbarOne['maximum'] = len(Barcode_value)
        cou = 0
        for x in Barcode_value:
            cou += 1
            sht_name = 'DVDQ-'+str(x)
            sheet_names_feiNone = [x.strip() for x in sheet_names]
            if sht_name in sheet_names_feiNone:
                index = sheet_names_feiNone.index(sht_name)
                sht = self.wb_RPT.sheets[index]
            else:
                sht_names = self.wb_RPT.sheet_names
                sht = self.wb_RPT.sheets.add(sht_name, after=sht_names[-1])

            sht.range('a1').value = 'DVDQ数据'
            for i in range(len(Cycle_list)):
                if sht[0:1, 1+8*i:2+8*i].value is not None:
                    continue
                sht[0:1, 1+8*i:2+8*i].value = Cycle_list[i]
                sht[1:2, 1+8*i:2+8*i].value = '充电容量(Ah)'
                sht[1:2, 2+8*i:3+8*i].value = '电压（V)'
                sht[1:2, 3+8*i:4+8*i].value = 'DV/DQ'
                sht[1:2, 4+8*i:5+8*i].value = 'DQ/DV'
                sht[1:2, 5+8*i:6+8*i].value = '放电容量(Ah)'
                sht[1:2, 6+8*i:7+8*i].value = '电压（V)'
                sht[1:2, 7+8*i:8+8*i].value = 'DV/DQ'
                sht[1:2, 8+8*i:9+8*i].value = 'DQ/DV'

                # 设置格式
                sht[0:1, 1 + 8 * i:9 + 8 * i].api.Merge()
                sht[0:1, 1 + 8 * i:9 + 8 * i].api.HorizontalAlignment = -4108
                sht[0:1, 1 + 8 * i:9 + 8 * i].api.VerticalAlignment = -4108
                sht[0:, 1 + 8 * i:9 + 8 * i].api.Borders(7).LineStyle = 1
                sht[0:, 1 + 8 * i:9 + 8 * i].api.Borders(7).Weight = 3
                sht[1:2, 1+8*i:3+8*i].api.Font.Bold = True
                sht[1:2, 1+8*i:3+8*i].column_width = 13.5
                sht[1:2, 1 + 8 * i:5 + 8 * i].color = (75, 172, 198)
                sht[1:2, 5+8*i:9+8*i].api.Font.Bold = True
                sht[1:2, 5+8*i:9+8*i].column_width = 13.5
                sht[1:2, 5+8*i:9+8*i].color = (155, 187, 89)

            self.progressbarOne['value'] = cou
            self.root.update()
            if cou == len(Barcode_value):
                tkinter.messagebox.showinfo(message='完成')
                self.progressbarOne['value'] = 0

        app_data.quit()


Cycle_Suspended_test_course()
