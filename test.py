import sys
from PySide2.QtWidgets import (QWidget, QLabel, QHBoxLayout,
                             QComboBox, QApplication)

class Example(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        hbox = QHBoxLayout()

        self.lbl = QLabel()
        self.lbl.setPixmap('1.jpg')

        self.combo = QComboBox()
        self.combo.addItem('1.jpg')
        self.combo.addItem('2.jpg')
        self.combo.activated.connect(self.onActivated)

        hbox.addWidget(self.combo)
        hbox.addWidget(self.lbl)
        self.setLayout(hbox)

        self.move(300, 300)
        self.setWindowTitle('QComboBox')
        self.show()

    def onActivated(self, ind):
        self.lbl.setPixmap(self.combo.itemText(ind))



if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())