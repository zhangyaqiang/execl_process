# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'widget.ui'
#
# Created by: PyQt5 UI code generator 5.7
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
import xlrd,xlwt

class Ui_Widget(object):
    def setupUi(self, Widget):
        Widget.setObjectName("Widget")
        Widget.resize(400, 300)
        self.label = QtWidgets.QLabel(Widget)
        self.label.setGeometry(QtCore.QRect(70, 100, 81, 20))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Widget)
        self.label_2.setGeometry(QtCore.QRect(70, 150, 81, 20))
        self.label_2.setObjectName("label_2")
        self.confirm_button = QtWidgets.QPushButton(Widget)
        self.confirm_button.setGeometry(QtCore.QRect(140, 210, 113, 32))
        self.confirm_button.setObjectName("confirm_button")
        self.confirm_button.clicked.connect(self.convert)
        self.src_file = QtWidgets.QTextEdit(Widget)
        self.src_file.setGeometry(QtCore.QRect(180, 90, 171, 31))
        self.src_file.setObjectName("src_file")
        self.des_file = QtWidgets.QTextEdit(Widget)
        self.des_file.setGeometry(QtCore.QRect(180, 140, 171, 31))
        self.des_file.setObjectName("des_file")

        self.retranslateUi(Widget)
        QtCore.QMetaObject.connectSlotsByName(Widget)

    def retranslateUi(self, Widget):
        _translate = QtCore.QCoreApplication.translate
        Widget.setWindowTitle(_translate("Widget", "Widget"))
        self.label.setText(_translate("Widget", "源文件地址"))
        self.label_2.setText(_translate("Widget", "目的文件地址"))
        self.confirm_button.setText(_translate("Widget", "确定"))

    def convert(self):
        src_file = str(self.src_file.toPlainText())
        print(src_file)
        des_file = str(self.des_file.toPlainText())
        des_file += '.xls'
        print(des_file)
        src_file = src_file.replace('\\', '/')
        des_file = des_file.replace('\\', '/')
        print(src_file)
        print(des_file)
        self.data_process(src_file, des_file)

    def data_process(self,src_file, des_file):
        src_data = []
        des_data = []
        data = xlrd.open_workbook(src_file)
        print(data)
        table = data.sheet_by_index(0)
        nrows = table.nrows
        for i in range(0, nrows):
            src_data.append(table.row_values(i))
        des_data.append(src_data[0])
        del src_data[0]
        i = 1
        j = len(src_data)-1
        src_company = src_data[0][0]
        des_company = src_data[0][1]
        while i != len(src_data)-1:
            if i == j:
                src_company = src_data[i][0]
                des_company = src_data[i][1]
                j = len(src_data)-1
                continue
            row = src_data[i]
            if (row[0] == src_company and row[1] == des_company) or (row[0] == des_company and row[1] == src_company):
                i += 1
                continue
            row = src_data[j]
            if (row[0] == src_company and row[1] == des_company) or (row[0] == des_company and row[1] == src_company):
                temp = src_data[i]
                src_data[i] = row
                src_data[j] = temp
                i += 1
                j -= 1
            else:
                j -= 1
        des_data += src_data
        wbk = xlwt.Workbook(encoding='ascii')
        sheet = wbk.add_sheet('sheet 1')
        for i in range(0, len(des_data)):
            for j in range(0, len(des_data[i])):
                sheet.write(i, j, des_data[i][j])
        wbk.save(des_file)


if __name__=="__main__":
    import sys

    app=QtWidgets.QApplication(sys.argv)
    widget=QtWidgets.QWidget()
    ui=Ui_Widget()
    ui.setupUi(widget)
    widget.show()
    sys.exit(app.exec_())
