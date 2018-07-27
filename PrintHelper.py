# -*- encoding: UTF-8 -*-

import sys
from PyQt5.QtWidgets import QApplication, QWidget, QGridLayout, QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog, QDesktopWidget
from PyQt5.QtGui import QIcon, QTextCursor
from PyQt5.QtCore import pyqtSlot, pyqtSignal, QObject, QThread, QTimer

import time
import os
import os.path

import xlrd

import win32print
import win32api

reload(sys)
sys.setdefaultencoding('utf-8')

class EmittingStream(QObject):
	textWritten = pyqtSignal(str)
	def write(self, text):
		self.textWritten.emit(str(text))	

class App(QWidget):
	def __init__(self):
		super(App, self).__init__()
		#初始化变量
		self.root = None
		self.excel = None
		self.titie = "打印小助手"
		#用于设置窗体距屏幕左边的距离
		self.left = 20
		#用于设置窗体距屏幕上方的距离
		self.top = 20
		#用于设置窗体的宽度
		self.width = 640
		#用于设置窗体的高度
		self.height = 480
		self.initUI()
	def initUI(self):
		self.gridlayout = QGridLayout()
		#设置出20 x 20 的格局
		for i in range(20):
			self.gridlayout.setColumnStretch(i,1)
			self.gridlayout.setRowStretch(i,1)
		lbRoot = QLabel('查找的文件夹：')
		lbExcel = QLabel('打印文件的excel：')
		lbLog = QLabel('日志输入：')
		self.edRoot = QLineEdit()
		self.edExcel = QLineEdit()
		self.btnRoot = QPushButton("选择文件夹")
		self.btnRoot.setFixedSize(100,20)
		self.btnRoot.clicked.connect(self.clickRoot)
		self.btnExcel = QPushButton("选择excel文件")
		self.btnExcel.setFixedSize(100,20)
		self.btnExcel.clicked.connect(self.clickExcel)
		self.btnPrint = QPushButton("打印")
		self.btnPrint.setFixedSize(100,20)
		self.btnPrint.clicked.connect(self.clickPrint)
		self.edLog = QTextEdit()
		#添加控件
		self.gridlayout.addWidget(lbRoot, 0, 0)
		self.gridlayout.addWidget(lbExcel, 1, 0)
		self.gridlayout.addWidget(self.edRoot, 0, 1, 1, 15)
		self.gridlayout.addWidget(self.edExcel, 1, 1, 1, 15)
		self.gridlayout.addWidget(self.btnRoot, 0, 16, 1, 6)
		self.gridlayout.addWidget(self.btnExcel, 1, 16, 1, 6)
		self.gridlayout.addWidget(lbLog, 2, 0)
		self.gridlayout.addWidget(self.btnPrint, 2, 16, 1, 6)
		self.gridlayout.addWidget(self.edLog, 3, 0, 17, 20)
		#设置布局
		self.setLayout(self.gridlayout)
		#设置窗体的标题
		self.setWindowTitle(self.titie)
		#使用setGeometry(left, top, width, height)方法设置窗体的参数
		self.setGeometry(self.left, self.top, self.width, self.height)
		#通过调用show()函数来显示窗口
		self.show()
		#窗口居中
		screen = QDesktopWidget().screenGeometry()
		size = self.geometry()
		self.move((screen.width()-size.width())/2, (screen.height()-size.height())/2)
		#重定向输出
		sys.stdout = EmittingStream(textWritten=self.normalOutputWritten)
		sys.stderr = EmittingStream(textWritten=self.errorOutputWritten)
	def __del__(self):
		sys.stdout = sys.__stdout__
		sys.stderr = sys.__stderr__
	def normalOutputWritten(self, text):
		if (len(text.strip()) == 0):
			return
		cursor = self.edLog.textCursor()
		cursor.movePosition(QTextCursor.End)
		cursor.insertHtml('<font color="black">%s</font><br>' % text )
		self.edLog.setTextCursor(cursor)
		self.edLog.ensureCursorVisible()
	def errorOutputWritten(self, text):
		if (len(text.strip()) == 0):
			return
		cursor = self.edLog.textCursor()
		cursor.movePosition(QTextCursor.End)
		cursor.insertHtml('<font color="red">%s</font><br>' % text)
		self.edLog.setTextCursor(cursor)
		self.edLog.ensureCursorVisible()
	@pyqtSlot()
	def clickRoot(self):
		print '选择文件夹'
		self.root = QFileDialog.getExistingDirectory(self)
		print '选中文件夹：%s' % self.root
		self.edRoot.setText(self.root)
	@pyqtSlot()
	def clickExcel(self):
		print '选择excel文件'
		#QFileDialog.getOpenFileName
		self.excel,_ = QFileDialog.getOpenFileName(self, '选择excel文件', './', 'excel(*.xls *.xlsx)')
		print '选中excel文件：%s' % self.excel
		self.edExcel.setText(self.excel)
	@pyqtSlot()
	def clickPrint(self):
		print '开始打印'
		self.btnRoot.setEnabled(False)
		self.btnExcel.setEnabled(False)
		self.btnPrint.setEnabled(False)
		self.taskThread = TaskThread(self.root, self.excel)
		self.taskThread.taskFinished.connect(self.onTaskFinished)
		self.taskThread.start()
	def onTaskFinished(self):
		print '打印结束'
		self.btnRoot.setEnabled(True)
		self.btnExcel.setEnabled(True)
		self.btnPrint.setEnabled(True)

class TaskThread(QThread):
	taskFinished = pyqtSignal()
	def __init__(self, root, excel):
		super(TaskThread, self).__init__()
		self.root = root
		self.excel = excel
	def run(self):
		self.runTask()
		self.taskFinished.emit()
	def readExcel(self, file) :
		excelFile = xlrd.open_workbook(file)
		sheetName = excelFile.sheet_names()[0]
		sheet = excelFile.sheet_by_name(sheetName)
		
		result = []
		index = 0
		while index < sheet.nrows :
			item = sheet.cell(index,0).value.encode('utf-8')
			if (item != None and len(item.strip()) != 0) :
				result.append(item.strip())
			index = index + 1
		return result
	def findFile(self, root, file) :
		print '查找的文件夹： %s' % root
		for parent, dirnames, filenames in os.walk(root) :
			for filename in filenames :
				#print "parent is :" + parent
				#print "filename is:" + filename
				#print "the full name of the file is :" + os.path.join(parent, filename)
				if (filename == file) :
					fullname = os.path.join(parent, filename)
					#fullname = parent.replace('/') + '/' + filename
					return fullname
	def printFile(self, file) :
		if (not os.path.exists(file)) :
			sys.stderr.write('文件（%s）不存在\n' % file)
			return
			
		if win32print.GetDefaultPrinterW() == None :
			sys.stderr.write('找不到打印机')
			return
			
		win32api.ShellExecute(0,\
			'print',\
			file,\
			win32print.GetDefaultPrinterW(),\
			".",
			0)
	def runTask(self) :
		start = time.time()
		if (self.root == None or len(self.root.strip()) == 0):
			sys.stderr.write('请选择文件夹')
			return
		if (self.excel == None or len(self.excel.strip()) == 0):
			sys.stderr.write('请选择excel文件')
			return
		fileList = self.readExcel(self.excel)
		print 'excel中需要打印的文件：%s' % fileList
		
		success = 0
		fail = 0
		for file in fileList :
			target = self.findFile(self.root, file)
			if(target != None) :
				success = success + 1
				print '找到文件: %s' % target
				self.printFile(target)
			else :
				fail = fail + 1
				sys.stderr.write('找不到文件: %s\n' % file)
		end = time.time()
		print '打印时间：%d秒' % (end - start)
		print '打印结果：总数（%d），成功（%d），<font color="red">失败（%d）</font>。' % (len(fileList), success, fail)
		
def main() :
	app = QApplication(sys.argv)
	ex = App()
	app.exec_()


if __name__ == '__main__' :
	main()