import sys
import GenerateMaster
from PyQt5.QtWidgets import QMainWindow, QPushButton, QApplication


class Example(QMainWindow):

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):

        btnGenMast = QPushButton("Generate Master", self)
        btnGenMast.move(30, 50)

        btnGenMast.clicked.connect(self.genMastClicked)

        self.statusBar()

        self.setGeometry(300, 300, 290, 150)
        self.setWindowTitle('Event sender')
        self.show()

    def genMastClicked(self):
        GenerateMaster.main()


if __name__ == '__main__':

    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())
