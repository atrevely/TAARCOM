import sys
import pandas as pd
import os.path
from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import QMainWindow, QPushButton, QApplication, \
                            QFileDialog, QTextEdit, QComboBox, QLabel
from PyQt5.QtCore import pyqtSlot
import GenerateMaster
import MergeFixedEntries
import SalesReportGenerator


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

        # Custom output stream.
        sys.stdout = Stream(newText=self.onUpdateText)

        # Set working directory as current file directory.
        dirPath = os.path.dirname(os.path.realpath(__file__))
        os.chdir(dirPath)

        # Initialize the threadpool for handling worker jobs.
        self.threadpool = QtCore.QThreadPool()
        # Initialize UI and supporting filenames.
        self.initUI()
        self.filenames = []
        self.master = []

        # Show welcome message.
        print('Welcome to the TAARCOM Commissions Manager.\n'
              'Messages and updates will display below.\n'
              '______________________________________________________\n'
              'REMINDER: Did you pull the latest version from GitHub?\n'
              '---')

        # Initialize global variables.
        global fieldMappings
        lookDir = 'Z:/Commissions Lookup/'
        # Upload field mappings, if found.
        if os.path.exists(lookDir + 'fieldMappings.xlsx'):
            fieldMappings = pd.read_excel(lookDir + 'fieldMappings.xlsx',
                                          index_col=False)
        else:
            print('No field mappings found!\n'
                  'Please make sure fieldMappings.xlsx is in the directory.\n'
                  '***')

        # Try finding/loading the supporting files.
        if not os.path.exists(lookDir + 'Lookup Master - Current.xlsx'):
            print('No Lookup Master found!\n'
                  'Please make sure Lookup Master is in the directory.\n'
                  '***')
        if not os.path.exists(lookDir + 'distributorLookup.xlsx'):
            print('No distributor lookup found!\n'
                  'Please make sure distributorLookup.xlsx '
                  'is in the directory.\n'
                  '***')
        if not os.path.exists(lookDir + 'principalList.xlsx'):
            print('No principal list found!\n'
                  'Please make sure principalList.xlsx '
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
        # Check for existence of principal list file.
        princList = None
        lookDir = 'Z:/Commissions Lookup/'
        if os.path.exists(lookDir + 'principalList.xlsx'):
            # Load principal list.
            princList = pd.read_excel(lookDir + 'principalList.xlsx',
                                      index_col=False)

        # Button for generating the master list.
        self.btnGenMast = QPushButton('Process Raw Files\nto '
                                      'Running\nCommissions', self)
        self.btnGenMast.move(650, 400)
        self.btnGenMast.resize(150, 150)
        self.btnGenMast.clicked.connect(self.genMastClicked)
        self.btnGenMast.setToolTip('Process selected raw data files and '
                                   'append them to the selected Running '
                                   'Commissions.\nIf no Running Commissions '
                                   'is selected, starts a new one.')

        # Button for selecting files to compile into master list.
        self.btnOpenFiles = QPushButton('Select Raw\n Commission Files', self)
        self.btnOpenFiles.move(50, 30)
        self.btnOpenFiles.resize(150, 100)
        self.btnOpenFiles.clicked.connect(self.openFilesClicked)
        self.btnOpenFiles.setToolTip('Open explorer tool for selecting raw '
                                     'commission file(s) to process.')

        # Button for selecting a current master to append to.
        self.btnUploadMast = QPushButton('Select Running\n'
                                         'Commissions\nFile', self)
        self.btnUploadMast.move(250, 30)
        self.btnUploadMast.resize(150, 100)
        self.btnUploadMast.clicked.connect(self.uploadMastClicked)
        self.btnUploadMast.setToolTip('Open explorer tool for selecting a '
                                      'Running Commissions to use.')

        # Button for writing fixed entries.
        self.btnFixEntries = QPushButton('Copy Fixed ENF\nEntries to\n'
                                         'Running\nCommissions', self)
        self.btnFixEntries.move(850, 200)
        self.btnFixEntries.resize(150, 150)
        self.btnFixEntries.clicked.connect(self.fixEntriesClicked)
        self.btnFixEntries.setToolTip('Migrate finished lines in the Entries '
                                      'Need Fixing file over to the '
                                      'associated Running Commissions.\n'
                                      'A Running Commissions needs to be '
                                      'selected and will be matched to '
                                      'the Entries Needs Fixing by the '
                                      'date at the end of the filename.')

        # Button for generating sales reports.
        self.btnGenReports = QPushButton('Generate Reports\nand Migrate Data\n'
                                         'to Comm Master', self)
        self.btnGenReports.move(850, 400)
        self.btnGenReports.resize(150, 150)
        self.btnGenReports.clicked.connect(self.genReportsClicked)
        self.btnGenReports.setToolTip('Generate commission and revenue '
                                      'reports from a finished '
                                      'Running Commissions file,\nthen '
                                      'migrate the Running Commissions data '
                                      'over to the Commissions Master.')

        # Button for clearing filename and master choices.
        self.btnClearAll = QPushButton('Clear\nSelections', self)
        self.btnClearAll.move(450, 30)
        self.btnClearAll.resize(150, 100)
        self.btnClearAll.clicked.connect(self.clearAllClicked)
        self.btnClearAll.setToolTip('Clear all selected files from the '
                                    'workspace.')

        # Dropdown menu for selecting principal.
        self.princMenu = QComboBox(self)
        self.princMenu.resize(150, 30)
        self.princMenu.move(650, 100)
        # Fill in principals, if file is found.
        if princList is not None:
            self.princMenu.addItems(list(princList['Abbreviation']))
        else:
            print('Principal list not found!\n'
                  'Please make sure principalList.xlsx is '
                  'in the directory!')
        self.princLabel = QLabel('Select Principal:', self)
        self.princLabel.resize(150, 100)
        self.princLabel.move(650, 35)

        # Create the text output widget.
        self.textBox = QTextEdit(self, readOnly=True)
        self.textBox.ensureCursorVisible()
        self.textBox.setLineWrapColumnOrWidth(500)
        self.textBox.setLineWrapMode(QTextEdit.FixedPixelWidth)
        self.textBox.setFixedWidth(550)
        self.textBox.setFixedHeight(400)
        self.textBox.move(50, 150)

        # Set window size and title, then show the window.
        self.setGeometry(300, 300, 1100, 600)
        self.setWindowTitle('Commissions Manager 2.0')
        self.show()

    def genMastClicked(self):
        """Send the GenerateMaster execution to a worker thread."""
        worker = Worker(self.genMastExecute)
        self.threadpool.start(worker)

    def genReportsClicked(self):
        """Send the SalesReportGenerator execution to a worker thread."""
        worker = Worker(self.genReportsExecute)
        self.threadpool.start(worker)

    def fixEntriesClicked(self):
        """Send the MergeFixedEntries execution to a worker thread."""
        worker = Worker(self.fixEntriesExecute)
        self.threadpool.start(worker)

    def clearAllClicked(self):
        """Clear the filenames and master variables."""
        self.filenames = []
        self.master = []
        print('All file selections cleared.\n'
              '---')

    def genMastExecute(self):
        """Runs function for processing new files to master."""
        # Check to see if we're ready to process.
        princ = self.princMenu.currentText()
        lookDir = 'Z:/Commissions Lookup/'
        mapExists = os.path.exists(lookDir + 'fieldMappings.xlsx')
        if self.filenames and mapExists and princ != '(No Selection)':
            # Turn buttons off.
            self.lockButtons()
            # Run the GenerateMaster.py file.
            try:
                GenerateMaster.main(self.filenames, self.master,
                                    fieldMappings, princ)
            except Exception as error:
                print('Unexpected Python error:\n'
                      + str(error)
                      + '\nPlease contact your local coder.')
            # Clear the filename selections.
            self.filenames = []
            # Turn buttons back on.
            self.restoreButtons()
        elif not mapExists:
            print('File fieldMappings.xlsx not found!\n'
                  'Please check file location and try again.\n'
                  '---')

        elif not self.filenames:
            print('No commission files selected!\n'
                  'Use the Select Commission Files button to select files.\n'
                  '---')

        elif princ == '(No Selection)':
            print('Please select a principal from the dropdown menu!\n'
                  '---')

    def genReportsExecute(self):
        """Runs function for generating salesperson reports."""
        if self.master:
            # Turn buttons off.
            self.lockButtons()
            # Run the SalesReportGenerator.py file.
            try:
                SalesReportGenerator.main(self.master)
            except Exception as error:
                print('Unexpected Python error:\n'
                      + str(error)
                      + '\nPlease contact your local coder.')
            # Clear the master selection.
            self.master = []
            # Turn buttons back on.
            self.restoreButtons()
        else:
            print('No Running Commissions uploaded!\n'
                  '---')

    def fixEntriesExecute(self):
        """Copy over fixed entries to Master."""
        if self.master:
            # Turn buttons off.
            self.lockButtons()
            # Run the GenerateMaster.py file.
            try:
                MergeFixedEntries.main(self.master)
            except Exception as error:
                print('Unexpected Python error:\n'
                      + str(error)
                      + '\nPlease contact your local coder.')
            # Turn buttons back on.
            self.restoreButtons()
        else:
            print('Please upload the current Running Commissions file.\n'
                  '---')

    def uploadMastClicked(self):
        """Upload an existing master list."""
        # Grab an existing master list to append to.
        self.master, _ = QFileDialog.getOpenFileName(
                self, filter="Excel files (*.xls *.xlsx *.xlsm)")
        if self.master:
            print('Current master list provided:\n'
                  + self.master
                  + '\n---')
            if 'Running Commissions' not in self.master:
                print('Caution!\n'
                      'The file uploaded as a Running Commissions '
                      'does not appear to be correct.\n'
                      '---')

    def openFilesClicked(self):
        """Provide filepaths for new data to process using GenerateMaster."""

        # Let us know we're clearing old selections.
        if self.filenames:
            print('Selecting new files, old selections cleared...')

        # Grab the filenames to be passed into GenerateMaster.
        self.filenames, _ = QFileDialog.getOpenFileNames(
                self, filter="Excel files (*.xls *.xlsx *.xlsm)")

        # Check if the current master got uploaded as a new file.
        for name in self.filenames:
            if 'Running Master' in name:
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

    def lockButtons(self):
        self.btnGenMast.setEnabled(False)
        self.btnOpenFiles.setEnabled(False)
        self.btnUploadMast.setEnabled(False)
        self.btnClearAll.setEnabled(False)
        self.princMenu.setEnabled(False)
        self.btnFixEntries.setEnabled(False)
        self.btnGenReports.setEnabled(False)

    def restoreButtons(self):
        self.btnGenMast.setEnabled(True)
        self.btnOpenFiles.setEnabled(True)
        self.btnUploadMast.setEnabled(True)
        self.btnClearAll.setEnabled(True)
        self.princMenu.setEnabled(True)
        self.btnFixEntries.setEnabled(True)
        self.btnGenReports.setEnabled(True)


class Worker(QtCore.QRunnable):
    """Inherits from QRunnable to handle worker thread.

    param args -- Arguments to pass to the callback function.
    param kwargs -- Keywords to pass to the callback function.
    """
    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()
        # Store constructor arguments (re-used for processing).
        self.fn = fn
        self.args = args
        self.kwargs = kwargs

    @pyqtSlot()
    def run(self):
        """Initialise the runner function with passed args, kwargs."""
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
