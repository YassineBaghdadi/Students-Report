import time
from os.path import split

from PyQt5 import QtWidgets, QtGui, QtCore, uic
from PyQt5.QtChart import *
from PyQt5.Qt import Qt
from PyQt5.QtChart import QBarSet, QBarSeries, QChart, QBarCategoryAxis, QValueAxis, QChartView
from PyQt5.QtGui import QPainter

import sys, os, random, threading
import pandas as pd
import logging

from PyQt5.QtWidgets import QHeaderView, QTableWidgetItem


class Main(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        logging.info('app started')
        uic.loadUi(os.path.join(os.path.dirname(__file__), 'ui/home.ui'), self)
        self.browse.clicked.connect(self.get_path)
        self.paths = []
        self.proc.setEnabled(False)
        self.path_txt.currentTextChanged.connect(self.path_changed)
        self.loading = Loading()
        self.contents.addWidget(self.loading)
        self.proc.clicked.connect(self.start_proc)
        self.frame_2.setEnabled(False)



    def change_widget(self, wdget):
        self.clear_content()
        self.contents.addWidget(wdget)

    def start_proc(self):
        loading = Loading('proc.gif')
        self.clear_content()
        self.contents.addWidget(loading)
        self.df = pd.read_excel(self.file)
        print(self.df)

        start_year = []
        for i in self.df['Student ID']:
            if str(i)[0] == '1':
                start_year.append(int(f'14{str(i)[3:5]}'))
            elif str(i)[0] == '4':
                start_year.append(int(f'14{str(i)[1:3]}'))

        self.df['start_year'] = start_year

        self.df['year_in_college'] = self.df['Graduation Year'] - self.df['start_year']

        self.r1 =R1(self.df)
        self.clear_content()
        self.contents.addWidget(self.r1)
        self.frame_2.setEnabled(True)






    def path_changed(self):
        if self.path_txt.currentText() == '':
            self.path_label.setFixedWidth(0)
            self.proc.setEnabled(False)

        elif os.path.isfile(self.path_txt.currentText()):
            self.path_label.setPixmap(QtGui.QPixmap('src/approved.png'))
            self.path_label.setScaledContents(True)
            self.proc.setEnabled(True)
            self.path_label.setFixedWidth(38)
        else:
            self.path_label.setPixmap(QtGui.QPixmap('src/err.png'))
            self.path_label.setScaledContents(True)
            self.proc.setEnabled(False)
            self.path_label.setFixedWidth(38)

    def get_path(self):
        self.file = QtWidgets.QFileDialog.getOpenFileName(caption='Load File', filter="Excel (*.xlsx *.xls)")[0]
        if self.file:
            self.proc.setEnabled(True)
            self.paths.insert(0, self.file)
            self.path_txt.clear()
            self.path_txt.addItems(set(self.paths))
            self.path_txt.setCurrentText(self.paths[0])

    def clear_content(self):
        for i in reversed(range(self.contents.count())):
            self.contents.itemAt(i).widget().setParent(None)
        # self.content.addWidget(widget)


class R1(QtWidgets.QWidget):
    def __init__(self, df = None):
        super().__init__()
        uic.loadUi(os.path.join(os.path.dirname(__file__), 'ui/r1.ui'), self)
        self.df = df
        # self.table.itemSelectionChanged.connect(self.table_select_event)
        self.table.clear()
        self.table_header = ['Majors', '2 Years', '3 Years', '4 Years', '5 Yrs or more']
        self.table.setColumnCount(len(self.table_header))
        self.table.setHorizontalHeaderLabels(self.table_header)
        self.table.resizeColumnsToContents()

        for i in range(len(self.table_header)):
            self.table.horizontalHeader().setSectionResizeMode(i, QHeaderView.Stretch)

        majors = [i for i in set(list(self.df['Major']))]
        gk = df.groupby('Major')
        data = []
        for major in majors:
            _2y = 0
            _3y = 0
            _4y = 0
            _5ym = 0
            for yoc in gk.get_group(major)['year_in_college']:
                if yoc == 2:
                    _2y += 1
                elif yoc == 3:
                    _3y += 1
                elif yoc == 4:
                    _4y += 1
                elif yoc >= 5:
                    _5ym += 1



            # data.append({major : [_2y, _3y, _4y, _5ym]})
            data.append([major, _2y, _3y, _4y, _5ym])


        # self.table.setRowCount(len(data))
        #
        # self.table.setVerticalHeaderLabels(list(data.keys()))

        for r_n, r_d in enumerate(data):
            self.table.insertRow(r_n)
            for c_n, d in enumerate(r_d):
                self.table.setItem(r_n, c_n, QtWidgets.QTableWidgetItem(str(d)))

        # Graph(self.graph_layout)

        set0 = QBarSet('2 years')
        set1 = QBarSet('3 years')
        set2 = QBarSet('4 years')
        set3 = QBarSet('5 years or more')


        set0.append([i[1] for i in data])
        set1.append([i[2] for i in data])
        set2.append([i[3] for i in data])
        set3.append([i[4] for i in data])

        series = QBarSeries()
        series.append(set0)
        series.append(set1)
        series.append(set2)
        series.append(set3)

        chart = QChart()
        chart.addSeries(series)
        chart.setTitle('Bar Chart Demo')
        chart.setAnimationOptions(QChart.SeriesAnimations)

        axisX = QBarCategoryAxis()
        axisX.append(str(i) for i in range(1, len(data)+1))

        all_v = []
        for i in data:
            for x in i[1:]:
                all_v.append(x)
        axisY = QValueAxis()
        axisY.setRange(0, max(all_v))

        chart.addAxis(axisX, Qt.AlignBottom)
        chart.addAxis(axisY, Qt.AlignLeft)

        chart.legend().setVisible(True)
        chart.legend().setAlignment(Qt.AlignBottom)

        chartView = QChartView(chart)
        self.graph_layout.addWidget(chartView)

class Loading(QtWidgets.QWidget):
    def __init__(self, gif = 'loading.gif'):
        super().__init__()
        uic.loadUi(os.path.join(os.path.dirname(__file__), 'ui/loading.ui'), self)
        self.gif = QtGui.QMovie(f'src/{gif}')
        self.label.setMovie(self.gif)
        self.gif.start()


class Graph:
    def __init__(self, parent):
        set0 = QBarSet('X0')
        set1 = QBarSet('X1')
        set2 = QBarSet('X2')
        set3 = QBarSet('X3')
        set4 = QBarSet('X4')


        set0.append([random.randint(0, 10) for i in range(6)])
        set1.append([random.randint(0, 10) for i in range(6)])
        set2.append([random.randint(0, 10) for i in range(6)])
        set3.append([random.randint(0, 10) for i in range(6)])
        set4.append([random.randint(0, 10) for i in range(6)])
        
        series = QBarSeries()
        series.append(set0)
        series.append(set1)
        series.append(set2)
        series.append(set3)
        series.append(set4)
        
        chart = QChart()
        chart.addSeries(series)
        chart.setTitle('Bar Chart Demo')
        chart.setAnimationOptions(QChart.SeriesAnimations)
        
        months = ('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun')
        
        axisX = QBarCategoryAxis()
        axisX.append(months)
        
        axisY = QValueAxis()
        axisY.setRange(0, 15)
        
        chart.addAxis(axisX, Qt.AlignBottom)
        chart.addAxis(axisY, Qt.AlignLeft)
        
        chart.legend().setVisible(True)
        chart.legend().setAlignment(Qt.AlignBottom)
        
        chartView = QChartView(chart)
        
        parent.addWidget(chartView)

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    main = Main()
    main.show()
    sys.exit(app.exec_())
