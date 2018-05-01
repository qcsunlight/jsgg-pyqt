
"""
棉花灌溉量计算软件
版本：1.0.0
开发：西北农林科技大学
最后编辑 2018-04-28

开发框架信息：
Name: PyQt5
Version: 5.9.2
Summary: Python bindings for the Qt cross platform UI and application toolkit
Home-page: https://www.riverbankcomputing.com/software/pyqt/
Author: Riverbank Computing Limited
Author-email: info@riverbankcomputing.com
License: GPL v3
Requires: sip
Required-by:
依赖包信息：
Name: xlrd
Version: 1.1.0
Summary: Library for developers to extract data from Microsoft Excel (tm) spreadsheet files
Home-page: http://www.python-excel.org/
Author: John Machin
Author-email: sjmachin@lexicon.net
License: BSD
Requires:
Required-by:

Name: xlwt
Version: 1.3.0
Summary: Library to create spreadsheet files compatible with MS Excel 97/2000/XP/2003 XLS files, on any platform, with Python 2.6, 2.7, 3.3+
Home-page: http://www.python-excel.org/
Author: John Machin
Author-email: sjmachin@lexicon.net
License: BSD
Requires:
Required-by:
"""

import sys
from PyQt5.QtWidgets import (QMainWindow, QApplication, QDialog,
        QFileDialog, QMessageBox)
from PyQt5.uic import loadUi
from PyQt5.QtCore import (QFile)
import jsggWindow, aboutWindow, setWindow #JsggResource
import os, xlrd, xlwt, datetime
from PyQt5.QtSql import (QSqlQuery, QSqlDatabase)

"""
帮助窗口类
最后编辑：2018-04-28
显示软件相关信息
"""
class About(QDialog):

    def __init__(self):
        super().__init__()
        # loadUi('about.ui', self)
        # 加载窗口
        ui = aboutWindow.Ui_Dialog()
        ui.setupUi(self)
        #显示窗口
        self.show()
"""
灌溉目标默认值设置窗口
最后编辑：2018-04-28
设置灌溉目标默认值
"""
class Set(QDialog):

    def __init__(self):
        super().__init__()
        # self.ui = loadUi('editgoal.ui', self)
        # 加载窗口
        self.ui = setWindow.Ui_Dialog()
        self.ui.setupUi(self)
        #显示窗口
        self.show()

        #显示数据库中的默认值
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

    """
    更新按钮触发函数
    更新默认值
    """
    def slotUpdate(self):

        area = float(self.ui.lineEdit.text() or 0)
        depth = float(self.ui.lineEdit_2.text() or 0)
        goal = float(self.ui.lineEdit_3.text() or 0)
        expect = float(self.ui.lineEdit_4.text() or 0)
        #更新数据库
        query = QSqlQuery()
        query.exec('update sets set item = %s where id = 1' % str(area or 0))
        query.exec('update sets set item = %s where id = 2' % str(depth or 0))
        query.exec('update sets set item = %s where id = 3' % str(goal or 0))
        query.exec('update sets set item = %s where id = 4' % str(expect or 0))

        self.close()

    # 取消按钮触发函数，关闭窗口
    def slotCancel(self):
        self.close()


"""
主窗口
"""
class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.initUi()
    
    def initUi(self):
        # self.ui = loadUi('new.ui', self)
        self.ui = jsggWindow.Ui_MainWindow()
        self.ui.setupUi(self)
        self.show()
        # 设置按钮样式表
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
        #禁用尚不能使用的按钮
        self.ui.pushButtonNext.setEnabled(False)
        self.ui.pushButtonPrev.setEnabled(False)
        self.ui.pushButtonCal.setEnabled(False)
        self.ui.pushButtonCls.setEnabled(False)
        self.ui.pushButtonOut.setEnabled(False)
        #设置默认值
        self.dir_url = None
        self.out_url = None
        self.row_id = 1

    # 将数据库中的默认值填写到相应的位置
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

    # 导入按钮槽函数
    # 选择文件导入，将导入的数据显示在相应的位置
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
            # 数据导入，启用可以使用的按钮
            self.ui.pushButtonCal.setEnabled(True)
            self.ui.pushButtonCls.setEnabled(True)
            self.ui.pushButtonOut.setEnabled(True)
            self.ui.actionCal.setEnabled(True)
            self.ui.actionOutput.setEnabled(True)
            if self.nrows > 1:
                self.ui.pushButtonNext.setEnabled(True)
        except:
            pass

    # 显示数据函数，在主窗口显示数据
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

        # 计算平均水分和平均温度
        self.avrShui = round((float(res[2]) + float(res[4]) + float(res[6])) / 3, 3)
        self.avrWen = round((float(res[3]) + float(res[5]) + float(res[7])) / 3, 3)
        # print(self.avrShui, self.avrWen)
        self.ui.lineEdit_9.setText(str(self.avrShui))
        self.ui.lineEdit_10.setText(str(self.avrWen))

        self.checkLimit()
        #数据更新之后，进行计算操作，更新灌溉量
        self.slotCal()
        # 判断此时数据能否继续上下切换
        if self.row_id > 1:
            self.ui.pushButtonPrev.setEnabled(True)
        elif self.row_id == self.nrows:
            self.ui.pushButtonNext.setEnabled(False)
        elif self.row_id == 1:
            self.ui.pushButtonPrev.setEnabled(False)

    # 检测水分含量是否低于下限
    def checkLimit(self):
        self.limit = 10
        if self.avrShui < self.limit:
            # print(True)
            QMessageBox.warning(None,'警告！',
                        '当前土壤水分含量已经低于下限，请调整'
                             '灌溉目标水分含量！'
                        )
            self.ui.lineEdit_13.setFocus(True)


    # 计算槽函数，计算当前小区的灌溉水量
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

    # 清除数据槽函数，清空当前操作数据
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

    # 导出数据槽函数，将此次导入的数据，进行计算并导出
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

    # 显示上一条
    def slotPrev(self):
        self.row_id -= 1
        self.setData(self.file, self.row_id)

    # 显示下一条
    def slotNext(self):
        self.row_id += 1
        # print(self.row_id)
        self.setData(self.file, self.row_id)

    # 设置按钮触发函数
    # 设置Action槽函数
    def slotSet(self):
        set = Set()
        set.exec_()
        try:
            set.ui.pushButton.clicked.connect(self.setDefaults())
        except:
            pass
    # 关于的Action的槽函数
    def slotAbout(self):
        about = About()
        about.exec_()


# 初始化数据库
# 在窗口启动之前初始化
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
    # 创建主窗口
    ex = MainWindow()
    sys.exit(app.exec_())

