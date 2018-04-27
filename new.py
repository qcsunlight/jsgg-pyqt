import sys
from PyQt5.QtWidgets import (QMainWindow, QApplication, QDialog,
        QFileDialog, QMessageBox)
from PyQt5.uic import loadUi
from PyQt5.QtCore import (QFile)
import JsggResource
import os, xlrd, xlwt, datetime
from PyQt5.QtSql import (QSqlQuery, QSqlDatabase)


class About(QDialog):

    def __init__(self):
        super().__init__()
        loadUi('about.ui', self)
        self.show()

class Set(QDialog):

    def __init__(self):
        super().__init__()
        self.ui = loadUi('editgoal.ui', self)
        self.show()

        query = QSqlQuery('select item from sets')
        items = []
        while query.next():
            items.append(query.value(0))
        # print(items)
        if len(items) == 4:
            self.ui.lineEdit.setText(str(items[0]))
            self.ui.lineEdit_2.setText(str(items[1]))
            self.ui.lineEdit_3.setText(str(items[2]))
            self.ui.lineEdit_4.setText(str(items[3]))

    def slotUpdate(self):

        area = float(self.ui.lineEdit.text() or 0)
        depth = float(self.ui.lineEdit_2.text() or 0)
        goal = float(self.ui.lineEdit_3.text() or 0)
        expect = float(self.ui.lineEdit_4.text() or 0)
        query = QSqlQuery()
        query.exec('update sets set item = %s where id = 1' % str(area or 0))
        query.exec('update sets set item = %s where id = 2' % str(depth or 0))
        query.exec('update sets set item = %s where id = 3' % str(goal or 0))
        query.exec('update sets set item = %s where id = 4' % str(expect or 0))

        self.close()


    def slotCancel(self):
        self.close()



class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.initUi()
    
    def initUi(self):
        self.ui = loadUi('new.ui', self)
        self.show()
        self.ui.pushButtonPrev.setStyleSheet(
            "QPushButton{border-image: url(:/res/后退.png);}"
            "QPushButton:hover{border-image: url(:/res/backend.png);}"
            "QPushButton:pressed{border-image: url(:/res/backselected.png);}"
        )
        self.ui.pushButtonNext.setStyleSheet(
            "QPushButton{border-image: url(:/res/前进.png);}"
            "QPushButton:hover{border-image: url(:/res/frontend.png);}"
            "QPushButton:pressed{border-image: url(:/res/frontselected.png);}"
        )
        self.setDefaults()
        self.ui.pushButtonNext.setEnabled(False)
        self.ui.pushButtonPrev.setEnabled(False)
        self.ui.pushButtonCal.setEnabled(False)
        self.ui.pushButtonCls.setEnabled(False)
        self.ui.pushButtonOut.setEnabled(False)
        self.dir_url = None
        self.out_url = None
        self.row_id = 1

    def setDefaults(self):
        query = QSqlQuery('select item from sets')
        items = []
        while query.next():
           items.append(query.value(0))
        self.sets = items
        if len(items) == 4:
            self.ui.lineEdit_11.setText(str(items[0]))
            self.ui.lineEdit_12.setText(str(items[1]))
            self.ui.lineEdit_13.setText(str(items[2]))
            self.ui.lineEdit_14.setText(str(items[3]))

    def slotIn(self):
        try:
            dialog = QFileDialog()
            if self.dir_url is None:
                self.dir_url = dialog.directoryUrl().toString()
            fileName = dialog.getOpenFileName(self, '加载数据文件', directory=self.dir_url,
                                          filter='Excle Files (*.xls *.xlsx)')
            # print(fileName[0])
            self.file = fileName[0]
            self.setData(fileName[0], 1)
            self.ui.pushButtonCal.setEnabled(True)
            self.ui.pushButtonCls.setEnabled(True)
            self.ui.pushButtonOut.setEnabled(True)
            if self.nrows > 1:
                self.ui.pushButtonNext.setEnabled(True)
        except:
            pass

    def setData(self, fileName, row):
        data = xlrd.open_workbook(fileName)
        table = data.sheets()[0]
        self.nrows, ncols = table.nrows, table.ncols
        # print(nrows, ncols)
        res = []
        if self.nrows <= 1:
            return
        res.append(table.cell_value(row, 0))
        # self.row_id += 1
        for i in range(2, 9):
            res.append(table.cell_value(row, i))
        # print(res)

        self.ui.lineEdit_1.setText(str(res[0]))
        self.ui.lineEdit_2.setText(str(res[1]))
        self.ui.lineEdit_3.setText(str(res[2]))
        self.ui.lineEdit_4.setText(str(res[3]))
        self.ui.lineEdit_5.setText(str(res[4]))
        self.ui.lineEdit_6.setText(str(res[5]))
        self.ui.lineEdit_7.setText(str(res[6]))
        self.ui.lineEdit_8.setText(str(res[7]))

        self.avrShui = round((float(res[2]) + float(res[4]) + float(res[6])) / 3, 3)
        self.avrWen = round((float(res[3]) + float(res[5]) + float(res[7])) / 3, 3)
        # print(self.avrShui, self.avrWen)
        self.ui.lineEdit_9.setText(str(self.avrShui))
        self.ui.lineEdit_10.setText(str(self.avrWen))
        self.checkLimit()
        self.slotCal()
        if self.row_id > 1:
            self.ui.pushButtonPrev.setEnabled(True)
        elif self.row_id == self.nrows:
            self.ui.pushButtonNext.setEnabled(False)
        elif self.row_id == 1:
            self.ui.pushButtonPrev.setEnabled(False)

    def checkLimit(self):
        self.limit = 10
        if self.avrShui < self.limit:
            # print(True)
            QMessageBox.warning(None,'警告！',
                        '当前土壤水分含量已经低于下限，请调整'
                             '灌溉目标水分含量！'
                        )
            self.ui.lineEdit_13.setFocus(True)


    def slotCal(self):
        area = float(self.ui.lineEdit_11.text())
        depth = float(self.ui.lineEdit_12.text())
        goal = float(self.ui.lineEdit_13.text())
        expect = float(self.ui.lineEdit_14.text() or 0)
        if area==0 or depth==0 or goal==0:
            QMessageBox.warning(None, '警告！',
                                '请检查数据完整性！'
                                )
            self.ui.lineEdit_11.setFocus(True)
            return
        res = round(area * (goal - self.avrShui) * (depth - expect) / 1000, 2)
        self.ui.lineEdit.setText(str(res))

    def slotCls(self):
        self.ui.lineEdit.setText('')
        self.ui.lineEdit_1.setText('')
        self.ui.lineEdit_2.setText('')
        self.ui.lineEdit_3.setText('')
        self.ui.lineEdit_4.setText('')
        self.ui.lineEdit_5.setText('')
        self.ui.lineEdit_6.setText('')
        self.ui.lineEdit_7.setText('')
        self.ui.lineEdit_8.setText('')
        self.ui.lineEdit_9.setText('')
        self.ui.lineEdit_10.setText('')
        self.ui.pushButtonCal.setEnabled(False)
        self.ui.pushButtonCls.setEnabled(False)
        self.ui.pushButtonOut.setEnabled(False)
        self.ui.pushButtonNext.setEnabled(False)
        self.ui.pushButtonPrev.setEnabled(False)

    def slotOut(self):
        try:
            dialog = QFileDialog()
            # if self.out_url is None:
            #     self.out_url = dialog.directoryUrl().toString()
            fileName = dialog.getSaveFileName(self, '导出计算结果', directory=self.dir_url,
                                          filter='Excle Files (*.xls)')
            # print(fileName[0])
            outfile = fileName[0]

            indata = xlrd.open_workbook(self.file)
            insheet = indata.sheets()[0]
            self.nrows = insheet.nrows
            data = xlwt.Workbook(encoding='utf-8')
            sheet = data.add_sheet('sheet1')
            sheet.write(0, 0, label='操作时间：%s' % datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            sheet.write(1, 0, label='设备地址')
            sheet.write(1, 1, label='采集时间')
            sheet.write(1, 2, label='土壤平均水分(Vol%)')
            sheet.write(1, 3, label='小区面积(m^2)')
            sheet.write(1, 4, label='目标灌溉深度(mm)')
            sheet.write(1, 5, label='目标灌溉水分(Vol%)')
            sheet.write(1, 6, label='预计降水量(mm)')
            sheet.write(1, 7, label='灌溉量(m^3)')

            for i in range(1, self.nrows):
                obj = []
                obj.append(str(insheet.cell_value(i, 0)))
                obj.append(str(insheet.cell_value(i, 2)))
                avr = round((float(insheet.cell_value(i, 3))
                             + float(insheet.cell_value(i, 5))
                             + float(insheet.cell_value(1, 7))) / 3, 3)
                # print(type(avr))
                # print(avr)
                obj.append(str(avr))
                for item in self.sets:
                    obj.append(item)
                res = round(float(self.sets[0]) * (float(self.sets[2]) - avr)
                            * (float(self.sets[1]) - float(self.sets[3])) / 1000, 2)
                obj.append(str(res))
                for it in obj:
                    sheet.write(i+1, obj.index(it), label='%s' % str(it))

            data.save(outfile)
            QMessageBox.warning(None, '提示', '导出成功！')

        except Exception as e:
            print(e)
            pass

    def slotPrev(self):
        self.row_id -= 1
        self.setData(self.file, self.row_id)
        
    def slotNext(self):
        self.row_id += 1
        # print(self.row_id)
        self.setData(self.file, self.row_id)

    def slotSet(self):
        set = Set()
        set.exec_()
        try:
            set.ui.pushButton.clicked.connect(self.setDefaults())
        except:
            pass

    def slotAbout(self):
        about = About()
        about.exec_()



def initDb():
    dbFile = os.path.join(os.path.dirname(__file__), 'jsggdb.sqlite')
    create = not QFile.exists(dbFile)
    db = QSqlDatabase.addDatabase("QSQLITE")
    db.setDatabaseName(dbFile)
    if not db.open():
        QMessageBox.warning(None, '错误', '数据库发生错误，请尝试重新打开程序！')
        sys.exit(1)
    if create:
        query = QSqlQuery()
        query.exec('''
            CREATE TABLE sets (
              id INTEGER PRIMARY KEY UNIQUE NOT NULL,
              item TEXT NOT NULL)
        ''')
        query.exec('''
            INSERT INTO sets(id, item) VALUES (1, '100')
        ''')
        query.exec('''
            INSERT INTO sets(id, item) VALUES (2, '10')
        ''')
        query.exec('''
            INSERT INTO sets(id, item) VALUES (3, '40')
        ''')
        query.exec('''
            INSERT INTO sets(id, item) VALUES (4, '0')
        ''')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    initDb()
    ex = MainWindow()
    sys.exit(app.exec_())

