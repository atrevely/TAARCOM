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
        self.filenames = []
        self.master = []

        # Custom output stream.
        sys.stdout = Stream(newText=self.onUpdateText)

    def onUpdateText(self, text):
        # Print console output to text box widget.
        cursor = self.process.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.process.setTextCursor(cursor)
        self.process.ensureCursorVisible()

    def closeEvent(self, event):
        # Return stdout to defaults.
        sys.stdout = sys.__stdout__
        super().closeEvent(event)

    def initUI(self):

        # Button for generating the master list.
        btnGenMast = QPushButton("Process Files \n to Master", self)
        btnGenMast.move(450, 100)
        btnGenMast.resize(100, 100)

        # Button for selecting files to compile into master list.
        btnOpenFiles = QPushButton("Open Files", self)
        btnOpenFiles.move(30, 50)

        # Button for selecting files to compile into master list.
        btnUploadMast = QPushButton("Upload Master", self)
        btnUploadMast.move(150, 50)

        # Link the buttons to their function calls.
        btnGenMast.clicked.connect(self.genMastClicked)
        btnOpenFiles.clicked.connect(self.openFilesClicked)
        btnUploadMast.clicked.connect(self.uploadMastClicked)

        # Create the output widget.
        self.process = QTextEdit(self, readOnly=True)
        self.process.ensureCursorVisible()
        self.process.setLineWrapColumnOrWidth(500)
        self.process.setLineWrapMode(QTextEdit.FixedPixelWidth)
        self.process.setFixedWidth(400)
        self.process.setFixedHeight(150)
        self.process.move(30, 100)

        # Set window size and title, then show the window.
        self.setGeometry(300, 300, 600, 300)
        self.setWindowTitle('Generate Master')
        self.show()

    def genMastClicked(self):
        # Check to see if we've selected files to process.
        if self.filenames:
            # Run the GenerateMaster.py file.
            GenerateMaster.main(self.filenames, self.master)
        else:
            print('No new files selected!')
            print('Use the Open Files button to select files.')
            print('---')

    def uploadMastClicked(self):
        # Grab an existing master list to append to.
        self.master, _ = QFileDialog.getOpenFileName(self, filter="Excel files (*.xls *.xlsx *.xlsm)")
        if self.master:
            print('Master list provided:')
            print(self.master)
            print('---')
            if 'CurrentMaster' not in self.master:
                print('Caution!')
                print('The file uploaded as master does not appear to be a master list.')
                print('---')

    def openFilesClicked(self):
        # Grab the filenames to be passed into GenerateMaster.py
        self.filenames, _ = QFileDialog.getOpenFileNames(self, filter="Excel files (*.xls *.xlsx *.xlsm)")

        # Check if the current master got uploaded as a new file.
        for names in self.filenames:
            if 'CurrentMaster' in names:
                print('Master uploaded as new file.')
                print('Try uploading files again.')
                print('---')
                return

        # Print out the selected filenames.
        if self.filenames:
            print('Files selected:')
            for file in self.filenames:
                print(file)
            print('---')


if __name__ == '__main__':
    # Run the application.
    app = QApplication(sys.argv)
    app.aboutToQuit.connect(app.deleteLater)  # This may be a hack for Spyder, consider deleting after testing.
    gui = GenMast()
    sys.exit(app.exec_())
