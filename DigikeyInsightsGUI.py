import sys
import os.path
from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import QMainWindow, QPushButton, QApplication, \
                            QFileDialog, QTextEdit
from PyQt5.QtCore import pyqtSlot
import LookupSales
import CompileFeedback


class Stream(QtCore.QObject):
    """Redirects console output to text widget."""
    newText = QtCore.pyqtSignal(str)

    def write(self, text):
        self.newText.emit(str(text))

    # Pass the flush so we don't get an attribute error.
    def flush(self):
        pass


class GenMast(QMainWindow):
    """Main application window."""
    def __init__(self):
        super().__init__()

        # Initialize the threadpool for handling worker jobs.
        self.threadpool = QtCore.QThreadPool()
        # Initialize UI and supporting filenames.
        self.initUI()
        self.filename = []
        self.filenames = []

        # Custom output stream.
        sys.stdout = Stream(newText=self.onUpdateText)
        # Show welcome message.
        print('Welcome to the TAARCOM Digikey Insights Manager.\n'
              'Messages and updates will display below.\n'
              '______________________________________________________\n'
              'REMINDER: Did you pull the latest version from GitHub?\n'
              '---')

        # Try finding/loading the supporting files.
        lookDir = ''#'Z:/Commissions Lookup/'
        if not os.path.exists(lookDir + 'rootCustomerMappings.xlsx'):
            print('No Root Customer Mappings found!\n'
                  'Please make sure rootCustomerMappings'
                  'is in the directory.\n'
                  '***')

    def onUpdateText(self, text):
        """Write console output to text widget."""
        cursor = self.textBox.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.textBox.setTextCursor(cursor)
        self.textBox.ensureCursorVisible()

    def closeEvent(self, event):
        """Shuts down application on close."""
        # Return stdout to defaults.
        sys.stdout = sys.__stdout__
        super().closeEvent(event)

    def initUI(self):
        """Creates UI window on launch."""

        # Button for looking up salespeople in Insight Report.
        self.btnLookSales = QPushButton('Lookup \n Salespeople \n'
                                        'for \n Insight File', self)
        self.btnLookSales.move(650, 400)
        self.btnLookSales.resize(150, 150)
        self.btnLookSales.clicked.connect(self.lookSalesClicked)
        self.btnLookSales.setToolTip('Returns a Digikey LI file with '
                                     'salespeople looked up from Account List '
                                     'and rootCustomerMappings.')

        # Button for copying over comments.
        self.btnCompileFeedback = QPushButton('Compile \n Feedback', self)
        self.btnCompileFeedback.move(650, 200)
        self.btnCompileFeedback.resize(150, 150)
        self.btnCompileFeedback.clicked.connect(self.compileFeedbackClicked)
        self.btnCompileFeedback.setToolTip('Combine individual reports with '
                                           'feedback into one file, and '
                                           'append that file to Digikey '
                                           'Insight Master.')

        # Button for clearing selections.
        self.btnClearAll = QPushButton('Clear \n File Selections', self)
        self.btnClearAll.move(450, 30)
        self.btnClearAll.resize(150, 100)
        self.btnClearAll.clicked.connect(self.clearAllClicked)

        # Button for selecting new file to lookup salespeople.
        self.btnOpenInsight = QPushButton('Select New \n Digikey Insight \n'
                                          'File',
                                          self)
        self.btnOpenInsight.move(50, 30)
        self.btnOpenInsight.resize(150, 100)
        self.btnOpenInsight.clicked.connect(self.openInsightClicked)
        self.btnOpenInsight.setToolTip('Select a brand new Digikey LI file.')

        # Button for selecting files to append to master.
        self.btnOpenFinished = QPushButton('Select Files \n with Feedback',
                                           self)
        self.btnOpenFinished.move(250, 30)
        self.btnOpenFinished.resize(150, 100)
        self.btnOpenFinished.clicked.connect(self.openFinishedClicked)
        self.btnOpenFinished.setToolTip('Select a batch of finished files '
                                        'that have feedback from salespeople.')

        # Create the text output widget.
        self.textBox = QTextEdit(self, readOnly=True)
        self.textBox.ensureCursorVisible()
        self.textBox.setLineWrapColumnOrWidth(500)
        self.textBox.setLineWrapMode(QTextEdit.FixedPixelWidth)
        self.textBox.setFixedWidth(550)
        self.textBox.setFixedHeight(400)
        self.textBox.move(50, 150)

        # Set window size and title, then show the window.
        self.setGeometry(300, 300, 900, 600)
        self.setWindowTitle('Digikey Insights Manager 2.0')
        self.show()

    def lookSalesClicked(self):
        """Send the LookupSales execution to a worker thread."""
        worker = Worker(self.lookSalesExecute)
        self.threadpool.start(worker)

    def compileFeedbackClicked(self):
        """Send the WriteComments execution to a worker thread."""
        worker = Worker(self.compileFeedbackExecute)
        self.threadpool.start(worker)

    def clearAllClicked(self):
        """Clear the filename variables."""
        self.filenames = []
        self.filename = []
        print('All file selections cleared.\n'
              '---')
        self.restoreButtons()

    def compileFeedbackExecute(self):
        """Runs function for compiling feedback and appending to Master."""
        # Check to make sure we've selected files.
        if self.filenames:
            # Turn buttons off.
            self.lockButtons()
            # Run the GenerateMaster.py file.
            try:
                CompileFeedback.main(self.filenames)
            except Exception as error:
                print('Unexpected Python error:\n'
                      + str(error)
                      + '\nPlease contact your local coder.')
            # Clear files.
            self.filenames = []
            # Turn buttons back on.
            self.restoreButtons()

        elif not self.filename:
            print('No finished Insight files selected!\n'
                  'Use the Select Commission Files button to select files.\n'
                  '---')

    def lookSalesExecute(self):
        """Runs function for looking up salespeople."""
        # Check to see if we're ready to process.
        lookDir = ''#'Z:/Commissions Lookup/'
        mapExists = os.path.exists(lookDir + 'rootCustomerMappings.xlsx')
        if self.filename and mapExists:
            # Turn buttons off.
            self.lockButtons()
            # Run the GenerateMaster.py file.
            try:
                LookupSales.main(self.filename)
            except Exception as error:
                print('Unexpected Python error:\n'
                      + str(error)
                      + '\nPlease contact your local coder.')
            # Clear file.
            self.filename = []
            # Turn buttons back on.
            self.restoreButtons()

        elif not mapExists:
            print('File rootCustomerMappings.xlsx not found!\n'
                  'Please check file location and try again.\n'
                  '---')

        elif not self.filename:
            print('No Insight file selected!\n'
                  'Use the Select Commission Files button to select files.\n'
                  '---')

    def openInsightClicked(self):
        """Provide filepath for new data to process using LookupSales."""
        # Let us know we're clearing old selections.
        if self.filename:
            print('Selecting new file, old selection cleared...')

        # Grab the filenames to be passed into LookupSales.
        self.filename, _ = QFileDialog.getOpenFileName(
                self, filter="Excel files (*.xls *.xlsx *.xlsm)")

        # Check if the current master got uploaded as a new file.
        if 'Digikey Insight Master' in self.filename:
            print('Master uploaded as new file.\n'
                  'Try uploading files again.\n'
                  '---')
            return

        # Print out the selected filenames.
        if self.filename:
            print('File selected: ' + self.filename + '\n---')
            # Turn off/on the correct buttons.
            self.btnCompileFeedback.setEnabled(False)
            self.btnLookSales.setEnabled(True)

    def openFinishedClicked(self):
        """Provide filepath for new data to process using AppendInsights."""
        # Let us know we're clearing old selections.
        if self.filenames:
            print('Selecting new files, old selections cleared...')

        # Grab the filenames to be passed into AppendInsights.
        self.filenames, _ = QFileDialog.getOpenFileNames(
                self, filter="Excel files (*.xls *.xlsx *.xlsm)")

        # Check if the current master got uploaded as a new file.
        for name in self.filenames:
            if 'Digikey Insight Master' in name:
                print('Master uploaded as new file.\n'
                      'Try uploading files again.\n'
                      '---')
                return

        # Print out the selected filenames.
        if self.filenames:
            print('Files selected:')
            for file in self.filenames:
                print(file)
            print('---')
            # Turn off/on the correct buttons.
            self.btnCompileFeedback.setEnabled(True)
            self.btnLookSales.setEnabled(False)

    def lockButtons(self):
        self.btnOpenInsight.setEnabled(False)
        self.btnLookSales.setEnabled(False)
        self.btnOpenFinished.setEnabled(False)
        self.btnClearAll.setEnabled(False)
        self.btnCompileFeedback.setEnabled(False)

    def restoreButtons(self):
        self.btnOpenInsight.setEnabled(True)
        self.btnLookSales.setEnabled(True)
        self.btnOpenFinished.setEnabled(True)
        self.btnClearAll.setEnabled(True)
        self.btnCompileFeedback.setEnabled(True)


class Worker(QtCore.QRunnable):
    """Inherits from QRunnable to handle worker thread.

    param args -- Arguments to pass to the callback function.
    param kwargs -- Keywords to pass to the callback function.
    """
    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()
        # Store constructor arguments (re-used for processing)
        self.fn = fn
        self.args = args
        self.kwargs = kwargs

    @pyqtSlot()
    def run(self):
        """Initialize the runner function with passed args, kwargs."""
        self.fn(*self.args, **self.kwargs)


if __name__ == '__main__':
    # Run the application.
    app = QApplication(sys.argv)
    app.aboutToQuit.connect(app.deleteLater)
    # Font settings.
    font = QtGui.QFont()
    font.setPointSize(10)
    app.setFont(font)
    # Open main window.
    gui = GenMast()
    sys.exit(app.exec_())
