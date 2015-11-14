from PySide.QtGui import *
from PySide.QtCore import *
import MainWindow as window

class Form(QWidget, window.Ui_MainWindow):
	def __init__(self, parent = None):
		super(Form,self).__init__(parent)
		self.setupUi(self)						
								
app = QApplication([])
win = Form()
win.show()
app.exec_()