import sys
import GenerateMaster
from PyQt5.QtWidgets import QMainWindow, QPushButton, QApplication, QFileDialog, QTextEdit
from PyQt5 import QtCore, QtGui


class Stream(QtCore.QObject):
    newText = QtCore.pyqtSignal(str)

    def write(self, text):
        self.newText.emit(str(text))


class GenMast(QMainWindow):

    def __init__(self):
        super().__init__()

        self.initUI()

        # Custom output stream.
        sys.stdout = Stream(newText=self.onUpdateText)

    def onUpdateText(self, text):
        # Print to text box widget.
        cursor = self.process.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.process.setTextCursor(cursor)
        self.process.ensureCursorVisible()

    def __del__(self):
        # Return stdout to defaults.
        sys.stdout = sys.__stdout__

    def initUI(self):

        # Button for generating the master list.
        btnGenMast = QPushButton("Generate Master", self)
        btnGenMast.move(30, 50)

        # Button for selecting files to compile into master list.
        btnOpenFiles = QPushButton("Open Files", self)
        btnOpenFiles.move(150, 50)

        # Link the buttons to their function calls.
        btnGenMast.clicked.connect(self.genMastClicked)
        btnOpenFiles.clicked.connect(self.openFilesClicked)

        # Create the output widget.
        self.process = QTextEdit(self, readOnly=True)
        self.process.ensureCursorVisible()
        self.process.setLineWrapColumnOrWidth(500)
        self.process.setLineWrapMode(QTextEdit.FixedPixelWidth)
        self.process.setFixedWidth(400)
        self.process.setFixedHeight(150)
        self.process.move(30, 100)

        # Set window size and title, then show the window.
        self.setGeometry(300, 300, 500, 300)
        self.setWindowTitle('Generate Master')
        self.show()

    def genMastClicked(self):
        # Run the GenerateMaster.py file.
        GenerateMaster.main()

    def openFilesClicked(self):
        # Grab the filenames to be passed into GenerateMaster.py
        filenames = QFileDialog.getOpenFileNames(self)
        print(filenames)


if __name__ == '__main__':

    app = QApplication(sys.argv)
    gui = GenMast()
    sys.exit(app.exec_())
