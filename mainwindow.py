#! /usr/bin/env python3
# _*_ coding utf-8 _*_

"""
农作物灌溉决策系统主窗口 MainWindow
使用框架PyQt5.x
auther: qcsunlight
e-mail: shaopengyue@foxmail.com
last edited: 2018.04.08
"""

import sys
from PyQt5.QtWidgets import (QMainWindow, QApplication, QDialog,
        QTableView, QFileDialog, QMessageBox, QSystemTrayIcon,
        QAction, QMenu)
from PyQt5.uic import loadUi
from PyQt5.QtCore import (Qt, QFile, QVariant)
import JsggRes
import os
import xlrd
from PyQt5.QtGui import (QStandardItemModel, QStandardItem, QIcon)
from PyQt5.QtSql import (QSqlQuery, QSqlDatabase, QSqlTableModel)


class BootStrap(QDialog):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        loadUi('bootstrap.ui', self)
        self.setWindowFlag(Qt.FramelessWindowHint)
        self.show()


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        # self.style = '''
        #         QPushButton{background-color:grey;color:white;}
        #         #window{ background:pink; }
        #         #test{ background-color:black;color:white; }
        # '''
        # self.setStyleSheet(self.style)
        self.initUI()

    def setDefault(self, defaults):
        self.ui.lineEdit11.setText(str(defaults[0]))
        self.ui.lineEdit12.setText(str(defaults[1]))
        self.ui.lineEdit13.setText(str(defaults[2]))
        self.ui.lineEdit14.setText(str(defaults[3]))

    def addSystemTray(self):
        # set window tray
        minimizeAction = QAction("Mi&nimize", self, triggered=self.hide)
        maximizeAction = QAction("Ma&zimize", self, 
                triggered=self.showMaximized)
        restoreAction = QAction("&Restore", self, 
                triggered=self.showNormal)
        quitAction = QAction("&Quit", self, triggered=self.close)
        self.trayIconMenu = QMenu(self)
        self.trayIconMenu.addAction(minimizeAction)
        self.trayIconMenu.addAction(maximizeAction)
        self.trayIconMenu.addAction(restoreAction)
        self.trayIconMenu.addSeparator()
        self.trayIconMenu.addAction(quitAction)
        self.trayIcon = QSystemTrayIcon(self)
        self.trayIcon.setIcon(QIcon(":/res/tray1.png"))
        self.trayIcon.setContextMenu(self.trayIconMenu)
        self.trayIcon.show()

    def initUI(self):
        self.ui = loadUi('mainwindow.ui', self)
        # self.setWindowFlag(Qt.FramelessWindowHint)
        # show window
        # set window tray
        self.addSystemTray()
        self.show()
        self.defaults = getDefault()
        self.setDefault(self.defaults)

        # self.model = QStandardItemModel()
        # head = getHeader()
        # self.model.setHorizontalHeaderLabels(head)
        # self.ui.tableView.setModel(self.model)

    def closeEvent(self, event):
        if self.trayIcon.isVisible():
            self.trayIcon.hide()

    def actionQuitTriggered(self):
        self.close()

    def actionInputTriggered(self):
        try:
            fileName = QFileDialog.getOpenFileName(self, '加载数据文件',
                          './', 'Excle Files (*.xls *.xlsx)')
        # print(fileName[0])
            if fileName[0] is not None:
                data = getData(fileName[0])
                # print(data)
                self.ui.lineEdit1.setText(str(data[0]))
                self.ui.lineEdit2.setText(str(data[1]))
                self.ui.lineEdit3.setText(str(data[2]))
                self.ui.lineEdit4.setText(str(data[3]))
                self.ui.lineEdit5.setText(str(data[4]))
                self.ui.lineEdit6.setText(str(data[5]))
                self.ui.lineEdit7.setText(str(data[6]))
                self.ui.lineEdit8.setText(str(data[7]))
                self.avrShui = round((float(data[2]) + float(data[4]) + float(data[6])) / 3, 3)
                self.avrWen = round((float(data[3]) + float(data[5]) + float(data[7])) / 3, 3)
                self.ui.lineEdit9.setText(str(self.avrShui))
                self.ui.lineEdit10.setText(str(self.avrWen))
        except Exception as e:
            pass

    def actionEditTriggered(self):
        pass

    def actionSetTriggered(self):
        pass

    def actionOutputTriggered(self):
        pass

    def actionCalTriggered(self):
        self.ui.label_result.setText('灌溉量：')
        area = float(self.ui.lineEdit11.text() or 0)
        depth = float(self.ui.lineEdit12.text() or 0)
        goal = float(self.ui.lineEdit13.text() or 0)
        expect = float(self.ui.lineEdit14.text() or 0)
        res = round((depth - expect/1000) * (goal - self.avrShui) * area, 2)
        # print(res)
        text = self.ui.label_result.text()
        # print(text)
        text += str(res)
        text += ' m^3'
        self.ui.label_result.setText(text)
        pass

    def clearBtnClicked(self):
        self.ui.lineEdit1.setText('')
        self.ui.lineEdit2.setText('')
        self.ui.lineEdit3.setText('')
        self.ui.lineEdit4.setText('')
        self.ui.lineEdit5.setText('')
        self.ui.lineEdit6.setText('')
        self.ui.lineEdit7.setText('')
        self.ui.lineEdit8.setText('')
        self.ui.lineEdit9.setText('')
        self.ui.lineEdit10.setText('')
        self.ui.label_result.setText('灌溉量：')

    def updateValue(self, data):

        pass


def initDb():
    query = QSqlQuery()
    query.exec('''CREATE TABLE item (
                      id INTEGER PRIMARY KEY UNIQUE NOT NULL,
                      value TEXT NOT NULL)''')
    query.exec('PRAGMA FOREIGN_KEYS=ON')
    query.exec('''CREATE TABLE defaults (
                      did INTEGER PRIMARY KEY UNIQUE NOT NULL,
                      id INTEGER NOT NULL ,
                      value REAL NOT NULL ,
                      FOREIGN KEY (id) REFERENCES item(id))''')
    query.exec('insert into item (id, value) VALUES (01, \'小区标识\')')
    query.exec('insert into item (id, value) VALUES (11, \'土壤水分1(Vol%)\')')
    query.exec('insert into item (id, value) VALUES (12, \'土壤温度1(℃)\')')
    query.exec('insert into item (id, value) VALUES (13, \'土壤水分2(Vol%)\')')
    query.exec('insert into item (id, value) VALUES (14, \'土壤温度2(℃)\')')
    query.exec('insert into item (id, value) VALUES (15, \'土壤水分3(Vol%)\')')
    query.exec('insert into item (id, value) VALUES (16, \'土壤温度3(℃)\')')
    query.exec('insert into item (id, value) VALUES (21, \'平均水分(Vol%)\')')
    query.exec('insert into item (id, value) VALUES (22, \'平均温度(℃)\')')
    query.exec('INSERT into item (id, value) values (23, \'小区面积(m^2)\')')
    query.exec('INSERT into item (id, value) values (24, \'灌溉深度(m)\')')
    query.exec('INSERT into item (id, value) values (25, \'灌溉目标水分含量(Vol%)\')')
    query.exec('INSERT into item (id, value) values (26, \'预计降水量(mm)\')')
    query.exec('INSERT into item (id, value) values (27, \'灌溉水量(m^3)\')')
    query.exec('insert into defaults (did, id, value) VALUES (1, 23, 10)')
    query.exec('insert into defaults (did, id, value) VALUES (2, 24, 1)')
    query.exec('insert into defaults (did, id, value) VALUES (3, 25, 40)')
    query.exec('insert into defaults (did, id, value) VALUES (4, 26, 0)')

def getHeader():
    query = QSqlQuery('select value from item')
    header = []
    while query.next():
        # print(query.value(0))
        header.append(query.value(0))
    return header
def getData(fileName):
    data = xlrd.open_workbook(fileName)
    table = data.sheets()[0]
    nrows, ncols = table.nrows, table.ncols
    res = []
    if nrows <= 1:
        return res
    res.append(table.cell_value(1, 0))
    for i in range(2, 9):
        res.append(table.cell_value(1, i))
    # print(res)
    return res
def getDefault():
    query = QSqlQuery('select value from defaults')
    res = []
    while query.next():
        res.append(query.value(0))
    return res


if __name__ == '__main__':
    app = QApplication(sys.argv)
    dbFile = os.path.join(os.path.dirname(__file__), 'jsgg.sqlite')
    create = not QFile.exists(dbFile)
    db = QSqlDatabase.addDatabase("QSQLITE")
    db.setDatabaseName(dbFile)
    if not db.open():
        QMessageBox.warning(None, '错误', '数据库发生错误，请尝试重新打开程序！')
        sys.exit(1)
    if create:
        initDb()
    # ba = BootStrap()
    ex = MainWindow()
    sys.exit(app.exec_())