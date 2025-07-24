#!/usr/bin/python
# -*- coding: UTF-8 -*-

import xlwings as xw
from tkinter import *


class Chart:

    def __init__(self):
        self.wb = xw.books.active # 活动工作表
        self.sheet2 = self.wb.sheets['Summary'] # Summary页
        self.row = self.sheet2.range('d6').expand('table').rows.count - 1
        if self.row <= 12:
            self.row = 12
        self.height_0 = self.sheet2.range('A'+str(9+self.row)).top # 第一个表格距离顶端距离
        self.width_0 = 22 # 第一个表格距离左端距离
        self.chart1 = 0 # 图表初始化
        self.sheet_name = 'RPT数据' # 'DVDQ-203D9VA000419' # 工作表名称初始化
        self.Chart_name = "Discharge Capacity Retention VS Cycle NO.curve " # 表格标题初始化
        self.ChartX_name = "Cycle_No." # 表格X轴标题初始化
        self.ChartY_name = 'Cap Retention(%)' # 表格Y轴标题初始化
        self.width = 400 # 396.8503937008 14cm 图表宽度初始化
        self.height = 227 # 226.7716535433 8cm 图表高度初始化
        self.jiange = 23.249921798706055 # 表格间隔初始化
        if self.sheet2.range('B'+str(8+self.row)).value is None:
            self.sheet2.range('B'+str(8+self.row)).value = '图表Plot'
            self.sheet2.range('B'+str(8+self.row)).api.Font.Color = -65536
            self.sheet2.range('B'+str(8+self.row)).api.Font.Size = 12
            self.sheet2.range('B'+str(8+self.row)).api.Font.Bold = True
            self.sheet2.range('B'+str(8+self.row)).api.Font.Name = 'Times New Roman'
            self.sheet2.range('B'+str(8+self.row)).api.HorizontalAlignment = -4131
            self.sheet2.range('B'+str(8+self.row)).api.VerticalAlignment = -4107



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

    def Get_data(self, x):

        sht = self.wb.sheets[self.sheet_name]
        row_sht = sht.used_range.last_cell.row
        value_1 = sht[0:row_sht, 2:3].value

        title_index = [x for x in range(len(value_1)) if (x-1) % (34+30) == 0 and value_1[x] is not None]
        serious = len(title_index)+1

        Blank = self.sheet2.range(self.sheet2.range(f'A1:A{row_sht}'), self.sheet2.range(f'{self.changeNumToChar(serious)}1:{self.changeNumToChar(serious)}{row_sht}'))
        sheet_name = "='RPT数据'"
        zuhe_title = [sheet_name + f"!$c${x+1}" for x in title_index]
        zuhe_X = [sheet_name + f"!$c${title_index[x]+4}:$c${title_index[x]+33+30}" for x in range(len(title_index))]
        zuhe_Y_1 = [sheet_name + f"!$h${title_index[x]+4}:$h${title_index[x]+33+30}" for x in range(len(title_index))]
        zuhe_Y_2 = [sheet_name + f"!$f${title_index[x]+4}:$f${title_index[x]+33+30}" for x in range(len(title_index))]
        zuhe_Y_3 = [sheet_name + f"!$ae${title_index[x]+4}:$ae${title_index[x]+33+30}" for x in range(len(title_index))]
        zuhe_Y_4 = [sheet_name + f"!$ai${title_index[x]+4}:$ai${title_index[x]+33+30}" for x in range(len(title_index))]

        if x == 0:
            zuhe_Y = zuhe_Y_1
        elif x == 1:
            zuhe_Y = zuhe_Y_2
        elif x == 2:
            zuhe_Y = zuhe_Y_3
        else:
            zuhe_Y = zuhe_Y_4


        return Blank, zuhe_title, zuhe_X, zuhe_Y

    def chart_texts(self):
        chart1_name = [str(x).split('\'')[1] for x in self.sheet2.charts]
        if not chart1_name:
            return []
        else:
            chart_texts = [self.sheet2.charts(x).api[1].ChartTitle.Text for x in chart1_name]
            return chart_texts


    def draw_1(self, x):

        if x == 0:
            self.Chart_name = 'Discharge Capacity Retention VS Cycle NO.curve '
            self.ChartY_name = "Cap Retention(%)"
        elif x == 1:
            self.Chart_name = 'Discharge Capacity VS Cycle NO.curve'
            self.ChartY_name = "Capacity(Ah)"
        elif x == 2:
            self.Chart_name = 'DC-DCR VS Cycle NO.curve '
            self.ChartY_name = "DCR(mohm)"
        elif x == 3:
            self.Chart_name = 'CC-DCR VS Cycle NO.curve '
            self.ChartY_name = "DCR(mohm)"

        chart_texts = self.chart_texts()
        if self.Chart_name in chart_texts:
            return


        if x <= 1:
            self.sheet2.charts.add(left=self.width_0+x*(self.width+self.jiange), top=self.height_0+3, width=self.width, height=self.height)  # 指定位置添加表格 1110
        else:
            self.sheet2.charts.add(left=self.width_0 + (x-2) * (self.width + self.jiange), top=self.height_0 + self.height+self.jiange, width=self.width, height=self.height)  # 指定位置添加表格 1110
        chart1_name = [str(x).split('\'')[1] for x in self.sheet2.charts]

        self.chart1 = self.sheet2.charts(chart1_name[-1])  # 选中表格 self.sheet2.charts["图表 2"]  self.sheet2.charts(2)
        self.chart1.chart_type = 'xy_scatter_smooth'  # 设置图标类型是平滑线折线图  第一个标签为x值
        Blank, zuhe_title, zuhe_X, zuhe_Y = self.Get_data(x)
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
        # self.chart1.api[1].Legend.Height = 125  # 图例高度

        self.chart1.api[1].Axes(1).HasTitle = True # 添加X轴标题
        self.chart1.api[1].Axes(1).AxisTitle.Text = self.ChartX_name  # Axes(1) 设置x轴标题的名字  Axes(2)设置x轴标题的名字
        # self.chart1.api[1].Axes(1).TickMarkSpacing = 50  # x坐标轴轴线 标记间隔
        # self.chart1.api[1].Axes(1).TickLabelSpacing = 50  # x坐标轴轴线 标签间隔
        self.chart1.api[1].Axes(1).AxisTitle.Format.TextFrame2.TextRange.Font.NameComplexScript = 'Calibri (正文)' # 设置X轴标题字体
        self.chart1.api[1].Axes(1).AxisTitle.Format.TextFrame2.TextRange.Font.NameFarEast = 'Calibri (正文)'
        self.chart1.api[1].Axes(1).AxisTitle.Format.TextFrame2.TextRange.Font.Name = 'Calibri (正文)'
        self.chart1.api[1].Axes(1).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
        self.chart1.api[1].Axes(1).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = -1  # 0 取消加粗 -1 加粗

        self.chart1.api[1].Axes(2).HasTitle = True # 添加Y轴标题
        self.chart1.api[1].Axes(2).AxisTitle.Text = self.ChartY_name  # Axes(1) 设置x轴标题的名字  Axes(2)设置y轴标题的名字
        self.chart1.api[1].Axes(2).AxisTitle.Format.TextFrame2.TextRange.Font.NameComplexScript = 'Times New Roman' # 设置Y轴标题字体
        self.chart1.api[1].Axes(2).AxisTitle.Format.TextFrame2.TextRange.Font.NameFarEast = 'Times New Roman'
        self.chart1.api[1].Axes(2).AxisTitle.Format.TextFrame2.TextRange.Font.Name = 'Times New Roman'
        self.chart1.api[1].Axes(2).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 10.5
        self.chart1.api[1].Axes(2).AxisTitle.Format.TextFrame2.TextRange.Font.Bold = -1  # 0 取消加粗 -1 加粗
        # if x == 0:
        #     self.chart1.api[1].Axes(2).MaximumScale = 1  # y 坐标轴最大值


    def main(self, root1, progressbarOne):
        progressbarOne['maximum'] = 4
        for j in range(4):
            self.draw_1(j)
            progressbarOne['value'] = j + 1
            root1.update()

    def main_1(self):
        for j in range(4):
            self.draw_1(j)




if __name__ == '__main__':

    root = Tk()
    root.title('小程序')
    root.geometry('240x100')
    root.attributes('-topmost', True)

    label_1 = Label(root, text='请输入所画图表对应sheet页名称:')
    label_1.place(x=0, y=20)
    entry_1 = Entry(root)
    entry_1.place(x=0, y=60, width=200)

    btn1 = Button(root, text="运行", command=lambda: Chart().main_1())
    btn1.place(x=190, y=20)

    root.mainloop()

