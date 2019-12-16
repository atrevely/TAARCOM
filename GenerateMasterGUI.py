import sys
import pandas as pd
import os.path
from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import QMainWindow, QPushButton, QApplication, \
                            QFileDialog, QTextEdit
from PyQt5.QtCore import pyqtSlot
import GenerateMaster
import MergeFixedEntries
import SalesReportGenerator
import CommTools

VERSION = 'Development v2.3'
if os.path.exists('Z:\\Commissions Lookup'):
    lookDir = 'Z:\\Commissions Lookup'
else:
    lookDir = os.getcwd()

lookDir = 'Z:/Commissions Lookup/'


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
        print('Welcome to the TAARCOM Commissions Manager ' + VERSION + '!\n'
              'Messages and updates will display below.\n'
              '_______________________________________________________________'
              '\nREMINDER: Did you check for updates on GitHub?\n'
              'REMINDER: If new code was pulled or the branch changed, please '
              'close and relaunch the program.\n'
              '---')
        if not os.path.exists('Z:/'):
            print('No connection to Z:/ drive established! Working locally.'
                  '\n---')
        # Try loading/finding the supporting files.
        if os.path.exists(lookDir):
            if os.path.exists(lookDir + '\\fieldMappings.xlsx'):
                self.fieldMappings = pd.read_excel(lookDir + '\\fieldMappings.xlsx',
                                                   index_col=False)
            else:
                print('No field mappings found!\n'
                      'Please make sure fieldMappings.xlsx is in the directory.\n'
                      '***')
            if not os.path.exists(lookDir + '\\Lookup Master - Current.xlsx'):
                print('No Lookup Master found!\n'
                      'Please make sure Lookup Master is in the directory.\n'
                      '***')
            if not os.path.exists(lookDir + '\\distributorLookup.xlsx'):
                print('No distributor lookup found!\n'
                      'Please make sure distributorLookup.xlsx '
                      'is in the directory.\n'
                      '***')
        else:
            print('Could not find route to ' + lookDir + '\nPlease make sure you '
                  'are connected to the TAARCOM server, then relaunch the program.'
                  '\n***')

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
        # Button for generating the master list.
        self.btnGenMast = QPushButton('Process Raw Files\nto '
                                      'Running\nCommissions', self)
        self.btnGenMast.move(750, 430)
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
        self.btnFixEntries.move(750, 230)
        self.btnFixEntries.resize(150, 150)
        self.btnFixEntries.clicked.connect(self.fixEntriesClicked)
        self.btnFixEntries.setToolTip('Migrate finished lines in the Entries '
                                      'Need Fixing file over to the '
                                      'associated Running Commissions.\n'
                                      'A Running Commissions needs to be '
                                      'selected and will be matched to '
                                      'the Entries Needs Fixing by the '
                                      'date at the end of the filenames.')

        # Button for generating sales reports.
        self.btnGenReports = QPushButton('Run Reports\n(and Migrate Data\n'
                                         'to Comm Master,\nif applicable)',
                                         self)
        self.btnGenReports.move(750, 30)
        self.btnGenReports.resize(150, 150)
        self.btnGenReports.clicked.connect(self.genReportsClicked)
        self.btnGenReports.setToolTip('Generate commission and revenue '
                                      'reports,\nthen migrate the Running '
                                      'Commissions data over to the '
                                      'Commissions Master.'
                                      '\nIf no Running Commissions is '
                                      'provided, will run reports on most '
                                      'recent quarter of data in the '
                                      'Commissions Master.')

        # Button for clearing filename and master choices.
        self.btnClearAll = QPushButton('Clear\nSelections', self)
        self.btnClearAll.move(450, 30)
        self.btnClearAll.resize(150, 100)
        self.btnClearAll.clicked.connect(self.clearAllClicked)
        self.btnClearAll.setToolTip('Clear all selected files from the '
                                    'workspace.')

        # Button for updating Lookup Master.
        self.btnUpdateLookup = QPushButton('Update Lookup\nMaster from RC',
                                           self)
        self.btnUpdateLookup.move(50, 610)
        self.btnUpdateLookup.resize(180, 80)
        self.btnUpdateLookup.clicked.connect(self.updateLookupClicked)
        self.btnUpdateLookup.setToolTip('Update the Lookup Master using a\n'
                                        'Running Commissions file.')

        # Create the text output widget.
        self.textBox = QTextEdit(self, readOnly=True)
        self.textBox.ensureCursorVisible()
        self.textBox.setLineWrapColumnOrWidth(600)
        self.textBox.setLineWrapMode(QTextEdit.FixedPixelWidth)
        self.textBox.setFixedWidth(650)
        self.textBox.setFixedHeight(450)
        self.textBox.move(50, 150)

        # Set window size and title, then show the window.
        self.setGeometry(100, 100, 1000, 700)
        self.setWindowTitle('Commissions Manager ' + VERSION)
        self.show()

    def genMastClicked(self):
        """Send the GenerateMaster execution to a worker thread."""
        self.lockButtons()
        worker = Worker(self.genMastExecute)
        if self.threadpool.activeThreadCount() == 0:
            self.threadpool.start(worker)

    def genReportsClicked(self):
        """Send the SalesReportGenerator execution to a worker thread."""
        self.lockButtons()
        worker = Worker(self.genReportsExecute)
        if self.threadpool.activeThreadCount() == 0:
            self.threadpool.start(worker)

    def fixEntriesClicked(self):
        """Send the MergeFixedEntries execution to a worker thread."""
        self.lockButtons()
        worker = Worker(self.fixEntriesExecute)
        if self.threadpool.activeThreadCount() == 0:
            self.threadpool.start(worker)

    def updateLookupClicked(self):
        """Send the UpdateLookups execution to a worker thread."""
        self.lockButtons()
        worker = Worker(self.updateLookupExecute)
        if self.threadpool.activeThreadCount() == 0:
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
        mapExists = os.path.exists(lookDir + '\\fieldMappings.xlsx')
        if self.filenames and mapExists:
            # Run the GenerateMaster.py file.
            try:
                GenerateMaster.main(self.filenames, self.master, self.fieldMappings)
            except Exception as error:
                print('Unexpected Python error:\n'
                      + str(error)
                      + '\nPlease contact your local coder.')
            # Clear the filename selections.
            self.filenames = []
        elif not mapExists:
            print('File fieldMappings.xlsx not found!\n'
                  'Please check file location and try again.\n'
                  '---')
        elif not self.filenames:
            print('No commission files selected!\n'
                  'Use the Select Commission Files button to select files.\n'
                  '---')
        self.restoreButtons()

    def genReportsExecute(self):
        """Runs function for generating salesperson reports."""
        # Turn buttons off.
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

    def fixEntriesExecute(self):
        """Copy over fixed entries to Master."""
        if self.master:
            # Run the GenerateMaster.py file.
            try:
                MergeFixedEntries.main(self.master)
            except Exception as error:
                print('Unexpected Python error:\n'
                      + str(error)
                      + '\nPlease contact your local coder.')
        else:
            print('Please upload the current Running Commissions file.\n'
                  '---')
        self.restoreButtons()

    def updateLookupExecute(self):
        """Update the Lookup Master using a Running Commissions."""
        if self.master:
            # Run the GenerateMaster.py file.
            try:
                CommTools.extractLookups(self.master)
            except Exception as error:
                print('Unexpected Python error:\n'
                      + str(error)
                      + '\nPlease contact your local coder.')
        else:
            print('Please upload a Running Commissions file.\n'
                  '---')
        self.restoreButtons()

    def uploadMastClicked(self):
        """Upload an existing Running Commissions."""
        # Grab an existing Running Commissions to append to.
        self.master, _ = QFileDialog.getOpenFileName(
                self, filter="Excel files (*.xls *.xlsx *.xlsm)")
        if self.master:
            print('Current Running Commissions selected:\n'
                  + os.path.basename(self.master)
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
            if 'Running Commissions' in name:
                print('RC uploaded as raw file.\n'
                      'Try uploading files again.\n'
                      '---')
                return
        # Print out the selected filenames.
        if self.filenames:
            print('Files selected:')
            for file in self.filenames:
                print(os.path.basename(file))
            print('---')

    def lockButtons(self):
        self.btnGenMast.setEnabled(False)
        self.btnOpenFiles.setEnabled(False)
        self.btnUploadMast.setEnabled(False)
        self.btnClearAll.setEnabled(False)
        self.btnFixEntries.setEnabled(False)
        self.btnGenReports.setEnabled(False)
        self.btnUpdateLookup.setEnabled(False)

    def restoreButtons(self):
        self.btnGenMast.setEnabled(True)
        self.btnOpenFiles.setEnabled(True)
        self.btnUploadMast.setEnabled(True)
        self.btnClearAll.setEnabled(True)
        self.btnFixEntries.setEnabled(True)
        self.btnGenReports.setEnabled(True)
        self.btnUpdateLookup.setEnabled(True)


class Worker(QtCore.QRunnable):
    """Inherits from QRunnable to handle worker thread.

    fn -- Function for the worker to run.
    args -- Arguments to pass to the callback function.
    kwargs -- Keywords to pass to the callback function.
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
