#!/usr/bin/python
# -*- coding: UTF-8 -*-

import xlwings as xw
from tkinter import *
import RPT_Cycle
import tkinter.messagebox



class Chart:

    def __init__(self):
        self.wb = xw.books.active # 活动工作表
        self.sheet2 = self.wb.sheets['Summary'] # Summary页
        self.row = self.sheet2.range('d6').expand('table').rows.count - 1 # Summary页 barcode 数量
        if self.row == 1:
            self.sheet_name = [self.sheet2[6:6+self.row, 3:4].value] # Summary页 barcode
            self.sheet_name = ['DVDQ-'+x for x in self.sheet_name] # 由Summary页 barcode 组成的sheet页名称
        else:
            self.sheet_name = self.sheet2[6:6 + self.row, 3:4].value # Summary页 barcode
            self.sheet_name = ['DVDQ-' + x for x in self.sheet_name] # 由Summary页 barcode 组成的sheet页名称
        if self.row <= 12:
            self.row = 12
        self.height_0 = self.sheet2.range('A'+str(9+self.row)).top # 第一个表格距离顶端距离
        self.chart1 = 0 # 图表初始化
        self.Chart_name = [f'工况循环充电DV/DQ vs Capacity curve\n{x.split("-")[1]}' for x in self.sheet_name]# 表格标题初始化
        self.ChartX_name = "容量/Ah" # 表格X轴标题初始化
        self.ChartY_name = "DV\\DQ" # 表格Y轴标题初始化
        self.width = 400 # 396.8503937008 14cm 图表宽度初始化
        self.height = 227 # 226.7716535433 8cm 图表高度初始化
        self.jiange = 23.249921798706055 # 表格间隔初始化


    @staticmethod  # 数字转字符
    def changeNumToChar(s):
        a = [chr(i) for i in range(97, 123)]
        ss = ''
        b = []
        if s <= 26:
            b = [s]
        else:
            while s > 26:
                c = s // 26  # 商
                d = s % 26  # 余数
                b = b + [d]
                if d == 0:
                    s = c - 1
                else:
                    s = c
            b = b + [s]
        b.reverse()  # 反转
        for i in b:
            ss = ss + a[i - 1]
        return ss

    def exist_is_or_not(self, y): # 判断 sheet 页是否存在
        sht = 0
        try:
            sht = self.wb.sheets[self.sheet_name[y]]
        except Exception:
            sheet_names = self.wb.sheet_names
            for x in sheet_names:
                if self.sheet_name[y] in x:
                    try:
                        sht = self.wb.sheets[self.sheet_name[y]]
                    except Exception:
                        tkinter.messagebox.showinfo(message=self.sheet_name[y]+'    命名不规范或不存在')
                    break

        return sht

    def chart_texts(self): # 判断改图表是否存在
        chart1_name = [str(x).split('\'')[1] for x in self.sheet2.charts]
        if not chart1_name:
            return []
        else:
            chart_texts = [self.sheet2.charts(x).api[1].ChartTitle.Text for x in chart1_name]
            return chart_texts


    def Get_data(self, x, y):  # 图表所需数据源
        sht = self.exist_is_or_not(y)
        if sht == 0:
            return 0, 0, 0, 0

        row_sht = sht.used_range.last_cell.row
        col_sht = sht.used_range.last_cell.column
        value_1 = sht[0:1, 0:col_sht].value
        value_3 = sht[2:3, 0:col_sht].value
        value_1 = [value_1[x] if value_3[x] is not None else None for x in range(col_sht)]
        # title_index = [x for x in range(col_sht) if value_1[x] is not None]
        title_index = [1, 9, 17, 25, 33, 41, 49, 57, 65, 73, 81, 89, 97, 105, 113, 121, 129, 137, 145, 153]
        row_sht_list = [sht[1:2, x:x+1].expand('table').rows.count+1 for x in title_index]
        row_sht_list_Y = [sht[1:2, x+4:x+5].expand('table').rows.count+1 for x in title_index]

        serious = len(title_index)+1

        Blank = self.sheet2.range(self.sheet2.range(f'A1:A{row_sht}'), self.sheet2.range(f'{self.changeNumToChar(serious)}1:{self.changeNumToChar(serious)}{row_sht+serious}'))
        col_X_1 = [self.changeNumToChar(x + 1) for x in title_index]
        col_X_2 = [self.changeNumToChar(x + 5) for x in title_index]
        col_Y_1 = [self.changeNumToChar(x + 3) for x in title_index]
        col_Y_2 = [self.changeNumToChar(x + 4) for x in title_index]
        col_Y_3 = [self.changeNumToChar(x + 7) for x in title_index]
        col_Y_4 = [self.changeNumToChar(x + 8) for x in title_index]
        sheet_name = f"=\'{self.sheet_name[y]}\'"

        zuhe_title = [sheet_name + f"!${x}$1" for x in col_X_1]
        zuhe_X_1 = [sheet_name + f"!${col_X_1[x]}$3:${col_X_1[x]}${2000}" for x in range(len(col_X_1))]
        zuhe_Y_1 = [sheet_name + f"!${col_Y_1[x]}$3:${col_Y_1[x]}${2000}" for x in range(len(col_Y_1))]
        zuhe_Y_2 = [sheet_name + f"!${col_Y_2[x]}$3:${col_Y_2[x]}${2000}" for x in range(len(col_Y_2))]
        zuhe_X_2 = [sheet_name + f"!${col_X_2[x]}$3:${col_X_2[x]}${2000}" for x in range(len(col_X_2))]
        zuhe_Y_3 = [sheet_name + f"!${col_Y_3[x]}$3:${col_Y_3[x]}${2000}" for x in range(len(col_Y_3))]
        zuhe_Y_4 = [sheet_name + f"!${col_Y_4[x]}$3:${col_Y_4[x]}${2000}" for x in range(len(col_Y_4))]
        # zuhe_X_1 = [sheet_name + f"!${col_X_1[x]}$3:${col_X_1[x]}${row_sht_list[x]}" for x in range(len(col_X_1))]
        # zuhe_Y_1 = [sheet_name + f"!${col_Y_1[x]}$3:${col_Y_1[x]}${row_sht_list[x]}" for x in range(len(col_Y_1))]
        # zuhe_Y_2 = [sheet_name + f"!${col_Y_2[x]}$3:${col_Y_2[x]}${row_sht_list[x]}" for x in range(len(col_Y_2))]
        # zuhe_X_2 = [sheet_name + f"!${col_X_2[x]}$3:${col_X_2[x]}${row_sht_list_Y[x]}" for x in range(len(col_X_2))]
        # zuhe_Y_3 = [sheet_name + f"!${col_Y_3[x]}$3:${col_Y_3[x]}${row_sht_list_Y[x]}" for x in range(len(col_Y_3))]
        # zuhe_Y_4 = [sheet_name + f"!${col_Y_4[x]}$3:${col_Y_4[x]}${row_sht_list_Y[x]}" for x in range(len(col_Y_4))]
        if x == 0:
            zuhe_Y = zuhe_Y_1
            zuhe_X = zuhe_X_1
        elif x == 1:
            zuhe_Y = zuhe_Y_2
            zuhe_X = zuhe_X_1
        elif x == 2:
            zuhe_Y = zuhe_Y_3
            zuhe_X = zuhe_X_2
        else:
            zuhe_Y = zuhe_Y_4
            zuhe_X = zuhe_X_2

        return Blank, zuhe_title, zuhe_X, zuhe_Y

    def draw_1(self, x, y): # 图表绘制
        Blank, zuhe_title, zuhe_X, zuhe_Y = self.Get_data(x, y)
        if Blank == 0:
            return
        if x == 0:
            self.Chart_name = f'工况循环充电DV/DQ vs Capacity curve\n{self.sheet_name[y].split("-")[1]}'
            self.ChartY_name = "DV\\DQ"
        elif x == 1:
            self.Chart_name = f'工况循环充电DQ/DV vs Capacity curve\n{self.sheet_name[y].split("-")[1]}'
            self.ChartY_name = "DQ\\DV"
        elif x == 2:
            self.Chart_name = f'工况循环放电DV/DQ vs Capacity curve\n{self.sheet_name[y].split("-")[1]}'
            self.ChartY_name = "DV\\DQ"
        else:
            self.Chart_name = f'工况循环放电DQ/DV vs Capacity curve\n{self.sheet_name[y].split("-")[1]}'
            self.ChartY_name = "DQ\\DV"

        chart_texts = self.chart_texts()
        if self.Chart_name in chart_texts:
            return

        self.sheet2.charts.add(left=22+x*(self.width+self.jiange), top=self.height_0+(y+2)*(self.height+self.jiange), width=self.width, height=self.height)  # 指定位置添加表格 1110
        chart1_name = [str(x).split('\'')[1] for x in self.sheet2.charts]
        self.chart1 = self.sheet2.charts(chart1_name[-1])  # 选中表格 self.sheet2.charts["图表 2"]  self.sheet2.charts(2)
        self.chart1.chart_type = 'xy_scatter_smooth_no_markers'  # 设置图标类型是平滑线散点图  第一个标签为x值
        self.chart1.set_source_data(Blank)
        for i in range(len(zuhe_title)):
            self.chart1.api[1].SeriesCollection(i + 1).Name = zuhe_title[i]
            self.chart1.api[1].SeriesCollection(i + 1).XValues = zuhe_X[i]
            self.chart1.api[1].SeriesCollection(i + 1).Values = zuhe_Y[i]

        self.chart1.api[1].SetElement(2)  # 添加标题
        self.chart1.api[1].ChartTitle.Text = self.Chart_name  # 设置标题名称
        self.chart1.api[1].ChartTitle.Format.TextFrame2.TextRange.Font.NameComplexScript = "宋体 (标题)"
        self.chart1.api[1].ChartTitle.Format.TextFrame2.TextRange.Font.NameFarEast = "宋体 (标题)"
        self.chart1.api[1].ChartTitle.Format.TextFrame2.TextRange.Font.Name = "宋体 (标题)"
        self.chart1.api[1].ChartTitle.Format.TextFrame2.TextRange.Font.Size = 12
        self.chart1.api[1].ChartTitle.Format.TextFrame2.TextRange.Font.Bold = -1  # 0 取消加粗 -1 加粗
        self.chart1.api[1].ChartTitle.VerticalAlignment = -4108
        # 图表
        self.chart1.api[1].PlotArea.Format.Line.Visible = -1  # 有外边框
        self.chart1.api[1].PlotArea.Format.Line.ForeColor.ObjectThemeColor = 13  # 13 黑色
        self.chart1.api[1].PlotArea.Format.Line.ForeColor.Brightness = 0.5  # 颜色深浅
        self.chart1.api[1].PlotArea.Format.Line.Transparency = 0  # 透明度

        self.chart1.api[1].SetElement(330)  # 关闭垂直的网格线
        self.chart1.api[1].SetElement(328)  # 关闭垂直的网格线
        self.chart1.api[1].SetElement(101)  # 右侧显示图例

        self.chart1.api[1].Legend.IncludeInLayout = False  # 图例与图表重叠
        self.chart1.api[1].Legend.Format.Line.Visible = 1  # 图例有外边框
        self.chart1.api[1].Legend.Position = 2  # -4107	位于图表下方。/2	位于图表边框的右上角。/	-4161	位于自定义的位置上/-4131	位于图表的左侧。/-4152	位于图表的右侧/-4160	位于图表的上方。
        self.chart1.api[1].Legend.Font.Name = '宋体 (正文)'  # 字体类类型
        self.chart1.api[1].Legend.Font.Size = 10  # 字体大小
        self.chart1.api[1].Legend.Font.FontStyle = False  # 加粗
        self.chart1.api[1].Legend.Font.Italic = False  # 斜体
        self.chart1.api[1].Legend.Height = 125  # 图例高度

        self.chart1.api[1].Axes(1).HasTitle = True # 添加X轴标题
        self.chart1.api[1].Axes(1).AxisTitle.Text = self.ChartX_name  # Axes(1) 设置x轴标题的名字  Axes(2)设置x轴标题的名字
        # self.chart1.api[1].Axes(1).TickMarkSpacing = 50  # x坐标轴轴线 标记间隔
        # self.chart1.api[1].Axes(1).TickLabelSpacing = 50  # x坐标轴轴线 标签间隔
        self.chart1.api[1].Axes(1).AxisTitle.Format.TextFrame2.TextRange.Font.NameComplexScript = 'Calibri (正文)'
        self.chart1.api[1].Axes(1).AxisTitle.Format.TextFrame2.TextRange.Font.NameFarEast = 'Calibri (正文)'
        self.chart1.api[1].Axes(1).AxisTitle.Format.TextFrame2.TextRange.Font.Name = 'Calibri (正文)'
        self.chart1.api[1].Axes(1).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
        self.chart1.api[1].Axes(1).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = -1  # 0 取消加粗 -1 加粗

        self.chart1.api[1].Axes(2).HasTitle = True # 添加Y轴标题
        self.chart1.api[1].Axes(2).AxisTitle.Text = self.ChartY_name  # Axes(1) 设置x轴标题的名字  Axes(2)设置y轴标题的名字
        self.chart1.api[1].Axes(2).AxisTitle.Format.TextFrame2.TextRange.Font.NameComplexScript = 'Times New Roman'
        self.chart1.api[1].Axes(2).AxisTitle.Format.TextFrame2.TextRange.Font.NameFarEast = 'Times New Roman'
        self.chart1.api[1].Axes(2).AxisTitle.Format.TextFrame2.TextRange.Font.Name = 'Times New Roman'
        self.chart1.api[1].Axes(2).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 10.5
        self.chart1.api[1].Axes(2).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = -1  # 0 取消加粗 -1 加粗

    def main(self, root1, progressbarOne): # 主函数
        RPT_Cycle.Chart().main(root1, progressbarOne)
        progressbarOne['maximum'] = len(self.sheet_name)*4
        for y in range(len(self.sheet_name)):
            for j in range(4):
                self.draw_1(j, y)
                progressbarOne['value'] = y*4+j + 1
                root1.update()
                if y == len(self.sheet_name) - 1 and j == 3:
                    tkinter.messagebox.showinfo(message='完成')

    def main_1(self): # 主函数

        for y in range(len(self.sheet_name)):
            for j in range(4):
                self.draw_1(j, y)


    def clear(self): # 清除所有图表
        self.sheet2.range('B' + str(8 + self.row)).clear_contents()
        chart1_name = [str(x).split('\'')[1] for x in self.sheet2.charts]
        [self.sheet2.charts(x).delete() for x in chart1_name]




if __name__ == '__main__':

    root = Tk()
    root.title('工况_图表')
    root.geometry('180x80')
    root.attributes('-topmost', True)

    btn2 = Button(root, text="运行", command=lambda: Chart().main_1())
    btn2.place(x=40, y=20)
    btn3 = Button(root, text="清除", command=lambda: Chart().clear())
    btn3.place(x=100, y=20)

    root.mainloop()

