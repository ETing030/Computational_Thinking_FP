from PyQt5 import QtWidgets, QtGui, QtCore
from projectuitry import Ui_MainWindow
import sys
import pandas
import csv
import os
import xlrd
#import project

dir_path = os.path.dirname(os.path.realpath(__file__))
filename1 = os.path.join(dir_path, '20201203課程.xlsx')
filename2 = os.path.join(dir_path, '20201206 108學年下學期課程.xlsx')
filename3 = os.path.join(dir_path, '條件一覽 - 110學年轉系條件.csv')
filename4 = os.path.join(dir_path, '條件一覽 - 轉系統整.csv')
filename5 = os.path.join(dir_path, '條件一覽 - 109學年雙主條件.csv')
filename6 = os.path.join(dir_path, '條件一覽 - 雙主統整.csv')
filename7 = os.path.join(dir_path, '條件一覽 - 109學年輔系條件.csv')
filename8 = os.path.join(dir_path, '條件一覽 - 輔系統整.csv')


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # MainWindow Title
        self.setWindowTitle('Project :)')

        # StatusBar (狀態列)
        self.statusBar().showMessage('測試版本，仍有很多錯誤或仍需改善的地方 :D ')


        # Menu
        self.ui.retranslateUi(self)
        self.ui.actionExit_3.setShortcut('Ctrl+Q')    # 快捷鍵 action+自設名稱
        self.ui.actionExit_3.triggered.connect(app.exit) #app退出功能

        self.ui.actiontransfer.triggered.connect(self.open_url1)
        self.ui.actiondouble.triggered.connect(self.open_url2)

        # Hide (PushButton為例)
        self.ui.comboBox_2.hide()
        self.ui.comboBox_3.hide()
        # self.ui.comboBox_4.hide()
        # self.ui.comboBox_5.hide()
        self.ui.label_2.hide()
        self.ui.label_3.hide()
        self.ui.label_4.hide()
        # self.ui.label_5.hide()
        # self.ui.label_6.hide()
        self.ui.label_Link.hide()


# Label------
        # Label links
        self.ui.label_Link.setOpenExternalLinks(True)
        self.ui.label_Link.setText("<a href='https://www.hollandexam.com/hollandQuiz.aspx'>職涯職業興趣測驗(外部連結)</a>")

# ComboBox-----
        # ComboBox 下拉是選單 (選擇是否)
        choices_1 = ['是','否']
        self.ui.comboBox.addItems(choices_1)  # 把選項給定(加入)給comboBox
        self.ui.comboBox.currentIndexChanged.connect(self.display)  # 每次選項改變時呼叫self.display()的function
        self.display() # 參見display function

        # Event  參見showButtonEvent function
        self.ui.comboBox.currentTextChanged.connect(self.showComboBoxEvent)
        #self.ui.comboBox_2.connect(self.showEvent_1)

        # ComboBox_2 下拉是選單 (否:選擇類型)
        choices_2= ['實做型','研究型','藝術型','社交型','企業型','常規型']
        self.ui.comboBox_2.addItems(choices_2)  # 把選項給定(加入)給comboBox
        self.ui.comboBox_2.currentIndexChanged.connect(self.display)  # 每次選項改變時呼叫self.display()的function
        #self.display() # 參見display function

        # ComboBox_3 下拉式選單 (否:選擇學院)
        self.ui.comboBox_2.currentTextChanged.connect(self.updateComboBox_3)
        self.ui.comboBox_3.currentIndexChanged.connect(self.display)  # 每次選項改變時呼叫self.display()的function
        #self.display() # 參見display function

        # ComboBox_4 下拉是選單 (是:選擇學院)
        self.ui.comboBox.currentTextChanged.connect(self.updateComboBox_4)
        self.ui.comboBox_4.currentIndexChanged.connect(self.display)  # 每次選項改變時呼叫self.display()的function

        # ComboBox_5 下拉是選單 (選擇科系)
        self.ui.comboBox.currentTextChanged.connect(self.updateComboBox_5_1)
        self.ui.comboBox_5.currentIndexChanged.connect(self.display)

        # ComboBox_6 (採取方式)
        choices_6 = ['轉系','雙主修','輔系','自由選修']
        self.ui.comboBox_6.addItems(choices_6)
        self.ui.comboBox_6.currentIndexChanged.connect(self.display)

        # PushButton
        if self.ui.comboBox_6.currentText()!='' and (self.ui.comboBox_5.currentText()!='請選擇' or self.ui.comboBox_5.currentText()!=''):
            self.ui.pushButton.clicked.connect(self.buttonClicked)  # 按鈕連接到點擊的函式

        # checkBox
        self.ui.checkBox.stateChanged.connect(self.IsChecked)

# table-----
        self.ui.table_course_2.setRowCount(0)
        self.ui.table_course_2.setColumnCount(0)
        self.ui.table_course_3.setRowCount(0)
        self.ui.table_course_3.setColumnCount(0)

        # tabWidget
        self.ui.tabWidget.addTab(self.ui.table_course_3,'下學期')




# function-----
    def open_url1(self):
        url = QtCore.QUrl('learnpyqt.com/tutorials/actions-toolbars-menus/')


    def open_url2(self):
        url = QtCore.QUrl('https://nckustory.ncku.edu.tw/formapply/index.php?auth')


    def IsChecked(self,state):
        if state==QtCore.Qt.Checked:
            self.ui.table_transfer.showRow(0)
            self.ui.table_double.showRow(0)
            #print('Checked')
        else:
            self.ui.table_transfer.hideRow(0)
            self.ui.table_double.hideRow(0)
            #print('Unchecked')



    def buttonClicked(self):
        self.ui.table_course.clear()
        self.ui.table_course_2.clear()
        self.ui.table_course_3.clear()
        self.ui.table_course.setRowCount(0)
        self.ui.table_course.setColumnCount(0)
        self.ui.table_course_2.setRowCount(1)
        self.ui.table_course_2.setColumnCount(7)
        self.ui.table_course_3.setRowCount(1)
        self.ui.table_course_3.setColumnCount(7)
        self.ui.label_10.clear()
        text = self.ui.comboBox_5.currentText()
        text2 = self.ui.comboBox_6.currentText()
        Department = {
            '1': '機械工程學系', '2': '化學工程學系', '3': '材料科學及工程學系', '4': '資源工程學系', '5': '土木工程學系', '6': '水利及海洋工程學系',
            '7': '系統及船舶機電工程學系', '8': '能源工程', '9': '航空太空工程學系',
            '10': '工程科學系', '11': '環境工程學系', '12': '測量及空間資訊學系', '13': '生物醫學工程學系', '14': '建築學系', '15': '都巿計劃學系',
            '16': '工業設計學系', '17': '電機工程學系',
            '18': '資訊工程學系', '19': '數學系', '20': '物理學系', '21': '化學系', '22': '光電科學與工程學系', '23': '地球科學系', '24': '生命科學系',
            '25': '生物科技與產業科學系', '26': '護理學系', '27': '醫學系',
            '28': '醫學檢驗生物技術學系', '29': '物理治療學系', '30': '職能治療學系', '31': '藥學系', '32': '牙醫學系', '33': '中國文學系',
            '34': '外國語文學系', '35': '台灣文學系', '36': '歷史學系', '37': '法律學系', '38': '政治學系',
            '39': '經濟學系', '40': '心理學系', '41': '企業管理學系', '42': '統計學系', '43': '會計學系', '44': '工業與資訊管理學系', '45': '交通管理科學系',
            '46': '不分系'
        }
        Department2 = {
            '1': '機械工程學系', '2': '化工系', '3': '材料系', '4': '資源系', '5': '土木工程學系', '6': '水利及海洋工程學系',
            '7': '系統系', '8': '能源學程', '9': '航空太空工程學系',
            '10': '工科系', '11': '環境工程學系', '12': '測量系', '13': '醫工系', '14': '建築學系', '15': '都巿計劃學系',
            '16': '工業設計學系', '17': '電機工程學系',
            '18': '資訊系', '19': '數學系', '20': '物理學系', '21': '化學系', '22': '光電系', '23': '地球科學系', '24': '生命科學系',
            '25': '生技系', '26': '護理學系', '27': '醫學系',
            '28': '醫技系', '29': '物治系', '30': '職能治療學系', '31': '藥學系', '32': '牙醫系', '33': '中國文學系',
            '34': '外國語文學系', '35': '台灣文學系', '36': '歷史學系', '37': '法律學系', '38': '政治學系',
            '39': '經濟學系', '40': '心理學系', '41': '企業管理學系', '42': '統計系', '43': '會計學系', '44': '工資系', '45': '交通管理科學系',
            '46': '不分系'
        }

        list2 = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期天']
        data = [0, 1, 2, 3, 4, 5, 6]
        data2 = [0, 1, 2, 3, 4, 5, 6]
        for i, j in enumerate(list2):
            data[i] = pandas.read_excel('20201203課程.xlsx', sheet_name=j, engine='openpyxl')
            data2[i] = pandas.read_excel('20201206 108學年下學期課程.xlsx', sheet_name=j, engine='openpyxl')

        # 上學期課程
        k = [0, 0, 0, 0, 0, 0, 0]
        # rows=data[1].shape[0]
        for i in range(len(list2)):
            for row in range(data[i].shape[0]):
                # print(row)
                if len([s for s in str(data[i].iat[row, 0])[0:3] if s in Department2[(text.split('.')[0])]]) == 3 and str(data[i].iat[row, 0]) != '':
                    for t in str(data[i].iat[row,0])[0]:                                # 他該死的系統系和統計系XDD 害我要加這兩行
                        if t == str(Department2[(text.split('.')[0])])[0]:
                            # print(str(data[i].iat[row,0])[0:3])
                            if data[i].iat[row, 4] != '':
                                if self.ui.table_course_2.rowCount() - 1 < k[i]:
                                    self.ui.table_course_2.insertRow(self.ui.table_course_2.rowCount())
                                    # print(self.ui.table_course2.rowCount())
                                self.ui.table_course_2.setItem(k[i], i, QtWidgets.QTableWidgetItem(data[i].iat[row, 4]))
                                k[i] = k[i] + 1

        # 下學期課程
        k = [0, 0, 0, 0, 0, 0, 0]
        for i in range(len(list2)):
            for row in range(data2[i].shape[0]):
                # print(row)
                if len([s for s in str(data2[i].iat[row, 0])[0:3] if s in Department2[(text.split('.')[0])]]) == 3 and str(data2[i].iat[row, 0]) != '':
                    for t in str(data2[i].iat[row,0])[0]:                                # 他該死的系統系和統計系XDD 害我要加這兩行
                        if t == str(Department2[(text.split('.')[0])])[0]:
                            # print(str(data2[i].iat[row,0])[0:3])
                            if data[i].iat[row, 4] != '':
                                if self.ui.table_course_3.rowCount() - 1 < k[i]:
                                    self.ui.table_course_3.insertRow(self.ui.table_course_3.rowCount())
                                    # print(self.ui.table_course3.rowCount())
                                self.ui.table_course_3.setItem(k[i], i, QtWidgets.QTableWidgetItem(data2[i].iat[row, 4]))
                                k[i] = k[i] + 1

        self.ui.table_course_2.setHorizontalHeaderLabels(['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日'])
        self.ui.table_course_3.setHorizontalHeaderLabels(['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日'])



        # 轉系
        if text2=='轉系':
            self.ui.tabWidget.setCurrentIndex(1)
            self.ui.table_transfer.clear()
            self.ui.table_transfer.setRowCount(1)
            self.ui.table_transfer.show()
            self.ui.table_double.hide()
            with open(filename3, encoding='utf-8') as csvfile:
                conditions=csv.DictReader(csvfile)
                for row in conditions:
                    if Department[(text.split('.')[0])] in row['院 系 別']:
                        self.ui.table_transfer.setItem(0, 0, QtWidgets.QTableWidgetItem('轉系條件\n'+row['轉 系 條 件']+'\n\n一般生可轉入名額：'+row['可轉入名額\n一般生']+'\n僑生可轉入名額：'+row['可轉入名額\n僑生']))

            with open(filename4, encoding='utf-8') as csvfile:
                conditions = csv.DictReader(csvfile)
                for row in conditions:
                    if Department[(text.split('.')[0])] in row['科系']:
                        #print(row['繳交資料'])
                        self.ui.table_transfer.insertRow(1)
                        self.ui.table_transfer.setItem(1, 0, QtWidgets.QTableWidgetItem(row['招生']))
                        if row['招生']=='':
                            self.ui.table_transfer.setItem(1, 0, QtWidgets.QTableWidgetItem('可招收轉系生'))

                        list = ['年級限制','成績要求','繳交資料','先修課程','測驗','面試','備註','心得']
                        for i,name in enumerate(list):
                            self.ui.table_transfer.insertRow(i+2)
                            self.ui.table_transfer.setItem(i+2,0,QtWidgets.QTableWidgetItem(row[name]))

                        self.ui.table_transfer.insertRow(i+3)
                        self.ui.table_transfer.setItem(i+3, 0, QtWidgets.QTableWidgetItem(row['一般生可轉入名額']))
                        self.ui.table_transfer.insertRow(i+4)
                        self.ui.table_transfer.setItem(i+4, 0, QtWidgets.QTableWidgetItem(row['橋生可轉入名額']))

                        self.ui.table_transfer.setVerticalHeaderLabels(['較原始資料','招收','年級限制','成績要求','繳交資料','先修課程','測驗','面試','備註','心得','一般生可轉入名額','僑生可轉入名額'])
                        self.ui.table_transfer.setHorizontalHeaderLabels([''])
                        row = self.ui.table_transfer.rowCount()
                        for i in range(row):
                            #print(self.ui.table_transfer.item(i,0).text())
                            if self.ui.table_transfer.item(i,0).text()=='':
                                self.ui.table_transfer.hideRow(i)


        # 雙主修
        elif text2=='雙主修':
            self.ui.tabWidget.setCurrentIndex(0)
            self.ui.table_double.clear()
            self.ui.table_double.setRowCount(1)
            self.ui.table_transfer.hide()
            self.ui.table_double.show()
            with open(filename5, encoding='utf-8') as csvfile:
                conditions=csv.DictReader(csvfile)
                for row in conditions:
                    if Department[(text.split('.')[0])] in row['院系別']:
                        self.ui.table_double.setItem(0, 0, QtWidgets.QTableWidgetItem('雙主修條件\n'+row['雙主修條件']))

            with open(filename6, encoding='utf-8') as csvfile:
                conditions = csv.DictReader(csvfile)
                for row in conditions:
                    if Department[(text.split('.')[0])] in row['科系']:
                        #print(row['繳交資料'])
                        self.ui.table_double.insertRow(1)
                        self.ui.table_double.setItem(1, 0, QtWidgets.QTableWidgetItem(row['招收']))
                        if row['招收']=='':
                            self.ui.table_double.setItem(1, 0, QtWidgets.QTableWidgetItem('可招收雙主修'))

                        list = ['年級限制','成績要求','繳交資料','先修課程','測驗','備註','聯絡電話']
                        for i,name in enumerate(list):
                            self.ui.table_double.insertRow(i+2)
                            self.ui.table_double.setItem(i+2,0,QtWidgets.QTableWidgetItem(row[name]))


                        # 課程備註
                        course_PS = row['課程備註']
                        course_PS = course_PS.strip('c(').rstrip(')')
                        course_PS = course_PS.replace('\"', '').split(',')
                        #print(course_PS)

                        if row['招收']=='' and row['課程備註']!='':
                            text3 = ''
                            for i,j in enumerate(course_PS):
                                #print(i+1,'. ',j.strip(' ').rstrip('。'))
                                a = '. '+j.strip(' ').rstrip('。')
                                text3 = text3+(str(i + 1))+a+'\n'
                            print(text3)
                            self.ui.label_10.setText(text3)

                        # 課程名稱
                        course = row['課程名稱']
                        course = course.strip('c(').rstrip(')')
                        course = course.replace('\"', '').split(',')
                        #print(course)
                        if row['招收']=='':
                            self.ui.table_course.setColumnCount(8)
                            self.ui.table_course.setHorizontalHeaderLabels(['課程名稱','學期','系號-序號','必選修','學分','時間','教室','備註'])
                            for i,j in enumerate(course):
                                self.ui.table_course.insertRow(i)
                                self.ui.table_course.setItem(i, 0, QtWidgets.QTableWidgetItem(j))
                                #print(i,j)
                                text4=j.replace(' ','').replace('（','(').replace('）',')').rstrip('*').rstrip('\\').rstrip('\'').rstrip('A').rstrip('B').rstrip('C').rstrip('D')
                                #print(text4)
                                self.course_info(i,list2,data,Department2,text,text4,'上')
                                self.course_info(i,list2,data2,Department2,text,text4,'下')



                        self.ui.table_double.setVerticalHeaderLabels(['較原始資料','招收','年級限制','成績要求','繳交資料','先修課程','測驗','備註','聯絡電話'])
                        self.ui.table_double.setHorizontalHeaderLabels([''])
                        row = self.ui.table_double.rowCount()
                        for i in range(row):
                            #print(self.ui.table_transfer.item(i,0).text())
                            if self.ui.table_double.item(i,0).text()=='':
                                self.ui.table_double.hideRow(i)


        # 輔系
        elif text2=='輔系':
            self.ui.tabWidget.setCurrentIndex(0)
            self.ui.table_double.clear()
            self.ui.table_double.setRowCount(1)
            self.ui.table_transfer.hide()
            self.ui.table_double.show()
            with open(filename7, encoding='utf-8') as csvfile:
                conditions=csv.DictReader(csvfile)
                for row in conditions:
                    if Department[(text.split('.')[0])] in row['院 系 別']:
                        self.ui.table_double.setItem(0, 0, QtWidgets.QTableWidgetItem('輔系條件\n'+row['輔 系 條 件']))

            with open(filename8, encoding='utf-8') as csvfile:
                conditions = csv.DictReader(csvfile)
                for row in conditions:
                    if Department[(text.split('.')[0])] in row['科系']:
                        #print(row['繳交資料'])
                        self.ui.table_double.insertRow(1)
                        self.ui.table_double.setItem(1, 0, QtWidgets.QTableWidgetItem(row['招收']))
                        if row['招收']=='':
                            self.ui.table_double.setItem(1, 0, QtWidgets.QTableWidgetItem('可招收輔系'))

                        list = ['年級限制','成績要求','繳交資料','先修課程','測驗','備註','聯絡電話']
                        for i,name in enumerate(list):
                            self.ui.table_double.insertRow(i+2)
                            self.ui.table_double.setItem(i+2,0,QtWidgets.QTableWidgetItem(row[name]))


                        # 課程備註
                        course_PS = row['課程備註']
                        course_PS = course_PS.strip('c(').rstrip(')')
                        course_PS = course_PS.replace('\"', '').split(',')
                        #print(course_PS)

                        if row['招收']=='' and row['課程備註']!='':
                            text3 = ''
                            for i,j in enumerate(course_PS):
                                #print(i+1,'. ',j.strip(' ').rstrip('。'))
                                a = '. '+j.strip(' ').rstrip('。')
                                text3 = text3+(str(i + 1))+a+'\n'
                            #print(text3)
                            self.ui.label_10.setText(text3)

                        # 課程名稱
                        course = row['課程名稱']
                        course = course.strip('c(').rstrip(')')
                        course = course.replace('\"', '').split(',')
                        #print(course)
                        if row['招收']=='':
                            self.ui.table_course.setColumnCount(8)
                            self.ui.table_course.setHorizontalHeaderLabels(['課程名稱','學期','系號-序號','必選修','學分','時間','教室','備註'])
                            for i,j in enumerate(course):
                                self.ui.table_course.insertRow(i)
                                self.ui.table_course.setItem(i, 0, QtWidgets.QTableWidgetItem(j))
                                #print(i,j)
                                text4=j.replace(' ','').replace('（','(').replace('）',')').rstrip('*').rstrip('\\').rstrip('\'')
                                #print(text4)
                                self.course_info(i,list2,data,Department2,text,text4,'上')
                                self.course_info(i,list2,data2,Department2,text,text4,'下')


                        self.ui.table_double.setVerticalHeaderLabels(['較原始資料','招收','年級限制','成績要求','繳交資料','先修課程','測驗','備註','聯絡電話'])
                        self.ui.table_double.setHorizontalHeaderLabels([''])
                        row = self.ui.table_double.rowCount()
                        for i in range(row):
                            #print(self.ui.table_transfer.item(i,0).text())
                            if self.ui.table_double.item(i,0).text()=='':
                                self.ui.table_double.hideRow(i)

        #else:
        #


        # 初始較原始資料顯示與否
        if self.ui.checkBox.checkState()!=QtCore.Qt.Checked:
            self.ui.table_transfer.hideRow(0)
            self.ui.table_double.hideRow(0)
        else:
            self.ui.table_transfer.showRow(0)
            self.ui.table_double.showRow(0)


        # 設定表格寬度和高度
        self.ui.table_transfer.setRowHeight(0,300)
        self.ui.table_transfer.setColumnWidth(0,450)
        self.ui.table_double.setRowHeight(0,300)
        self.ui.table_double.setColumnWidth(0,450)
        self.ui.table_course.setColumnWidth(0,150)
        self.ui.table_course.setColumnWidth(1,50)
        self.ui.table_course.setColumnWidth(2,80)
        self.ui.table_course.setColumnWidth(3,60)
        self.ui.table_course.setColumnWidth(4,50)
        self.ui.table_course.setColumnWidth(7,250)

        row1 = self.ui.table_transfer.rowCount()
        row2 = self.ui.table_double.rowCount()
        row3 = self.ui.table_course_2.rowCount()
        row4 = self.ui.table_course_3.rowCount()
        row5 = self.ui.table_course.rowCount()
        #column1 = self.ui.table_course.columnCount()
        column2 = self.ui.table_course_2.columnCount()
        for i in range(1,row1):
            self.ui.table_transfer.setRowHeight(i,50)
        for i in range(1,row2):
            self.ui.table_double.setRowHeight(i,50)
        for i in range(row3):
            self.ui.table_course_2.setRowHeight(i,50)
        for i in range(row4):
            self.ui.table_course_3.setRowHeight(i,50)
        for i in range(row5):
            self.ui.table_course.setRowHeight(i,50)
        #for i in range(column1):
        #    self.ui.table_course.setColumnWidth(i,150)
        for i in range(column2):
            self.ui.table_course_2.setColumnWidth(i,150)
            self.ui.table_course_3.setColumnWidth(i,150)

        # 文字置中(沒用)
        #self.ui.table_course.setTextElideMode(2)

        # 唯讀
        self.ui.table_transfer.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.ui.table_double.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.ui.table_course.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.ui.table_course_2.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.ui.table_course_3.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)



    def course_info(self,i,list2,data,Department2,text,text4,name): #self,i,list2,data,Department2,text,text4
        #print(text)
        for a in range(len(list2)):
            for row in range(data[a].shape[0]):
                if len([s for s in str(data[a].iat[row, 0])[0:3] if s in Department2[(text.split('.')[0])]]) == 3 and str(data[a].iat[row, 0]) != '':
                    for t in str(data[a].iat[row, 0])[0]:
                        if t == str(Department2[(text.split('.')[0])])[0]:
                             #print(str(data[i].iat[row,0])[0:3])
                            if data[a].iat[row, 4].split()[0].replace(' ', '').replace('（', '(').replace('）',')') == text4:
                                self.ui.table_course.setItem(i, 1, QtWidgets.QTableWidgetItem(name))
                                self.ui.table_course.setItem(i, 2, QtWidgets.QTableWidgetItem(data[a].iat[row, 1].split('\n')[0]))
                                self.ui.table_course.setItem(i, 3, QtWidgets.QTableWidgetItem(data[a].iat[row, 5].split()[1]))
                                self.ui.table_course.setItem(i, 4,QtWidgets.QTableWidgetItem(data[a].iat[row, 5].split()[0]))
                                self.ui.table_course.setItem(i, 5, QtWidgets.QTableWidgetItem(data[a].iat[row, 8].split('\n')[0].split()[0]))
                                if len(data[a].iat[row, 8].split('\n')) > 1:
                                    self.ui.table_course.setItem(i, 5, QtWidgets.QTableWidgetItem(data[a].iat[row, 8].split('\n')[0].split()[0] + '\n' + data[a].iat[row, 8].split('\n')[1].split()[0]))
                                if len(data[a].iat[row, 8].split('\n')[0].split())>2:
                                    self.ui.table_course.setItem(i,6,QtWidgets.QTableWidgetItem(data[a].iat[row, 8].split('\n')[0].split()[1] + '\n' + data[a].iat[row, 8].split('\n')[0].split()[2]))
                                elif len(data[a].iat[row, 8].split('\n')[0].split())==2:
                                    self.ui.table_course.setItem(i, 6, QtWidgets.QTableWidgetItem(data[a].iat[row, 8].split('\n')[0].split()[1]))

                                b = data[a].iat[row,4].split('\n')
                                b2 = []
                                #print(b)
                                for element in b:
                                    b2.append(element)
                                del b2[0]
                                #print(b2)
                                text5 = ''

                                if len(b[0].split(' '))>1:
                                    text5 += b[0].split(' ')[1]

                                for element in b2:
                                    if element!='' or element!=' ':
                                        text5 += element+'\n'
                                text5=text5.rstrip('\n')
                                #print(text5)
                                self.ui.table_course.setItem(i, 7, QtWidgets.QTableWidgetItem(text5))

    def showComboBoxEvent(self,text):
        if text =="是":
            self.ui.comboBox_2.hide()
            self.ui.comboBox_3.hide()
            self.ui.label_2.hide()
            self.ui.label_3.hide()
            self.ui.label_4.hide()
            self.ui.label_Link.hide()

            self.ui.comboBox_4.show()
            self.ui.comboBox_5.show()
            self.ui.label_5.show()
            self.ui.label_6.show()

        elif text =="否":
            self.ui.comboBox_2.show()
            self.ui.comboBox_3.show()
            self.ui.comboBox_5.show()
            self.ui.label_2.show()
            self.ui.label_3.show()
            self.ui.label_4.show()
            self.ui.label_6.show()
            self.ui.label_Link.show()

            self.ui.comboBox_4.hide()
            self.ui.label_5.hide()

    # 針對ComboBox_2選擇要放入ComboBox_3的內容
    def updateComboBox_3(self, text):
        self.ui.comboBox_3.clear()
        if text == "實做型":
            self.ui.comboBox_3.addItems(('工學院','規設院','電資院'))
        elif text == "研究型":
            self.ui.comboBox_3.addItems(('理學院','生技院','醫學院'))
        elif text == "藝術型":
            self.ui.comboBox_3.addItems(('規設院','文學院'))
        elif text == "社交型":
            self.ui.comboBox_3.addItems(('社科院','醫學院'))
        elif text == "企業型":
            self.ui.comboBox_3.addItems(('管學院','社科院'))
        elif text == "常規型":
            self.ui.comboBox_3.addItems(('管學院','社科院'))

    # 針對ComboBox_1選擇要放入ComboBox_5的內容
    def updateComboBox_5_1(self,text):
        self.ui.comboBox_5.setCurrentIndex(0)
        if text=="是":
            self.ui.comboBox_4.currentTextChanged.connect(self.updateComboBox_5)
        if text=="否":
            self.ui.comboBox_3.currentTextChanged.connect(self.updateComboBox_5)

    # 針對ComboBox_3,4選擇要放入ComboBox_5的內容
    def updateComboBox_5(self, text):
        self.ui.comboBox_5.clear()
        if text == "工學院":
            self.ui.comboBox_5.addItems(('1.機械系', '2.化工系', '3.材料系', '4.資源系', '5.土木系', '6.水利系', '7.系統系', '8.能源學程', '9.航太系', '10.工科系', '11.環工系',
             '12.測量系', '13.醫工系'))
        elif text == "規設院":
            self.ui.comboBox_5.addItems(('14.建築系', '15.都計系', '16.工設系'))
        elif text == "電資院":
            self.ui.comboBox_5.addItems(('17.電機系', '18.資訊系'))
        elif text == "理學院":
            self.ui.comboBox_5.addItems(('19.數學系', '20.物理系', '21.化學系', '22.光電科學系', '23.地科系'))
        elif text == "生技院":
            self.ui.comboBox_5.addItems(('24.生科系', '25.生技系'))
        elif text == "醫學院":
            self.ui.comboBox_5.addItems(('26.護理系', '27.醫學系', '28.醫技系', '29.物治系', '30.職治系', '31.藥學系', '32.牙醫系'))
        elif text == "文學院":
            self.ui.comboBox_5.addItems(('33.中文系', '34.外文系', '35.台文系', '36.歷史系'))
        elif text == "社科院":
            self.ui.comboBox_5.addItems(('37.法律系', '38.政治系', '39.經濟系', '40.心理系'))
        elif text == "管學院":
            self.ui.comboBox_5.addItems(('41.企管系', '42.統計系', '43.會計系', '44.工資系', '45.交管系'))
        elif text == "不分系":
            self.ui.comboBox_5.addItems(('請選擇','46.不分系'))

    # 針對ComboBox選擇要放入ComboBox_4的內容
    def updateComboBox_4(self, text):
        self.ui.comboBox_4.setCurrentIndex(0)
        if text == "是":
            self.ui.comboBox_4.addItems(('工學院', '規設院', '電資院', '理學院', '生技院', '醫學院', '文學院', '社科院', '管學院','不分系'))

    # 把comboBox選擇結果顯示於label_3
    def display(self):
        self.ui.label_7.setText('請確認結果\n你所選的科系是：%s' % self.ui.comboBox_5.currentText() + '，辦法是：' + self.ui.comboBox_6.currentText())

    def exit(self):
        app.exit()



if __name__ == '__main__':
     app = QtWidgets.QApplication([])
     window = MainWindow()
     window.show()

     sys.exit(app.exec_())