import sys
import pandas as pd
import os.path
from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import QMainWindow, QPushButton, QApplication, \
                            QFileDialog, QTextEdit, QTreeWidget, \
                            QTreeWidgetItem, QInputDialog
from PyQt5.QtCore import pyqtSlot
import GenerateMaster


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
        self.filenames = []
        self.master = []

        # Custom output stream.
        sys.stdout = Stream(newText=self.onUpdateText)
        # Print welcome message.
        print('Welcome to the TAARCOM Commissions Manager.')
        print('Messages and updates will display here.')
        print('---')

        # Create a global varaiable for the field mappings.
        # We can now edit it in the ColumnEdit class, or just leave it alone.
        global fieldMappings
        # Upload field mappings, if found.
        if os.path.exists('fieldMappings.xlsx'):
            fieldMappings = pd.read_excel('fieldMappings.xlsx', index_col=False)
        else:
            print('No field mappings found!')
            print('Please make sure fieldMappings.xlsx is in the directory.')
            print('***')

        # Try finding the supporting files.
        if not os.path.exists('Lookup Master 6-22-18.xlsx'):
            print('No Lookup Master found!')
            print('Please make sure Lookup Master is in the directory.')
            print('***')
        if not os.path.exists('distributorLookup.xlsx'):
            print('No distributor lookup found!')
            print('Please make sure distributorLookup.xlsx is in the directory.')
            print('***')

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
        self.btnGenMast = QPushButton('Process Files \n to Running \n Commissions', self)
        self.btnGenMast.move(650, 400)
        self.btnGenMast.resize(150, 150)
        self.btnGenMast.clicked.connect(self.genMastClicked)

        # Button for selecting files to compile into master list.
        self.btnOpenFiles = QPushButton('Select \n Commission Files', self)
        self.btnOpenFiles.move(50, 30)
        self.btnOpenFiles.resize(150, 100)
        self.btnOpenFiles.clicked.connect(self.openFilesClicked)

        # Button for selecting a current master to append to.
        self.btnUploadMast = QPushButton('Select \n Running \n Commissions', self)
        self.btnUploadMast.move(250, 30)
        self.btnUploadMast.resize(150, 100)
        self.btnUploadMast.clicked.connect(self.uploadMastClicked)

        # Button for editing column names/tags.
        self.btnEditColumns = QPushButton('Edit Column Tags', self)
        self.btnEditColumns.move(450, 30)
        self.btnEditColumns.resize(150, 100)
        self.btnEditColumns.clicked.connect(self.editColumnsClicked)

        # Button for clearing filename and master choices.
        self.btnClearAll = QPushButton('Clear Filenames \n and Master', self)
        self.btnClearAll.move(650, 200)
        self.btnClearAll.resize(150, 150)
        self.btnClearAll.clicked.connect(self.clearAllClicked)

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
        self.setWindowTitle('Commissions Manager 2.0')
        self.show()

    def editColumnsClicked(self):
        """Opens new window for editing field mappings."""
        # Open new window with data tree and editing processes.
        if os.path.exists('fieldMappings.xlsx'):
            global fieldMappings
            fieldMappings = pd.read_excel('fieldMappings.xlsx', index_col=False)
            self.columnsWindow = ColumnEdit()
            self.columnsWindow.show()
            self.lockButtons()
        else:
            print('No field mappings file found!')
            print('Please make sure fieldMappings.xlsx is in the directory.')
            print('***')

    def genMastClicked(self):
        """Send the GenerateMaster execution to a worker thread."""
        worker = Worker(self.genMastExecute)
        self.threadpool.start(worker)

    def clearAllClicked(self):
        """Clear the filenames and master variables."""
        self.filenames = []
        self.master = []
        print('All files and master cleared.')
        print('---')

    def genMastExecute(self):
        """Runs function for processing new files to master."""
        # Check to see if we've selected files to process.
        if self.filenames and os.path.exists('fieldMappings.xlsx'):
            # Turn buttons off.
            self.lockButtons()
            # Run the GenerateMaster.py file.
            GenerateMaster.main(self.filenames, self.master, fieldMappings)
            # Turn buttons back on.
            self.restoreButtons()

        elif os.path.exists('fieldMappings.xlsx'):
            print('No commission files selected!')
            print('Use the Select Commission Files button to select files.')
            print('---')

    def uploadMastClicked(self):
        """Upload an existing master list."""
        # Grab an existing master list to append to.
        self.master, _ = QFileDialog.getOpenFileName(
                self, filter="Excel files (*.xls *.xlsx *.xlsm)")
        if self.master:
            print('Current master list provided:')
            print(self.master)
            print('---')
            if 'Running Master' not in self.master:
                print('Caution!')
                print('The file uploaded as master'
                      + 'does not appear to be a master list.')
                print('---')

    def openFilesClicked(self):
        """Provide filepaths for new data to process using GenerateMaster."""
        # Grab the filenames to be passed into GenerateMaster.py
        self.filenames, _ = QFileDialog.getOpenFileNames(
                self, filter="Excel files (*.xls *.xlsx *.xlsm)")

        # Check if the current master got uploaded as a new file.
        for names in self.filenames:
            if 'Running Master' in names:
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

    def lockButtons(self):
        self.btnGenMast.setEnabled(False)
        self.btnOpenFiles.setEnabled(False)
        self.btnUploadMast.setEnabled(False)
        self.btnEditColumns.setEnabled(False)
        self.btnClearAll.setEnabled(False)

    def restoreButtons(self):
        self.btnGenMast.setEnabled(True)
        self.btnOpenFiles.setEnabled(True)
        self.btnUploadMast.setEnabled(True)
        self.btnEditColumns.setEnabled(True)
        self.btnClearAll.setEnabled(True)


class Worker(QtCore.QRunnable):
    '''
    Inherits from QRunnable to handle worker thread.

    param args -- Arguments to pass to the callback function.
    param kwargs -- Keywords to pass to the callback function.
    '''

    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()
        # Store constructor arguments (re-used for processing)
        self.fn = fn
        self.args = args
        self.kwargs = kwargs

    @pyqtSlot()
    def run(self):
        '''
        Initialise the runner function with passed args, kwargs.
        '''
        self.fn(*self.args, **self.kwargs)


class ColumnEdit(QMainWindow):
    """Window for editing field mappings."""
    def __init__(self, parent=None):
        """Create UI for window on launch."""
        super().__init__()

        # Set window size and title.
        self.setGeometry(200, 200, 800, 550)
        self.setWindowTitle('Field Mappings')

        # Create the tree widget with column names.
        self.colTree = QTreeWidget(self)
        self.colTree.resize(600, 500)
        self.colTree.setColumnCount(1)
        self.colTree.setHeaderLabels(['TCOM Column Names'])

        # Create the button for adding data names.
        btnAddName = QPushButton('Add Principal \n Field Name', self)
        btnAddName.move(630, 10)
        btnAddName.resize(150, 100)
        btnAddName.clicked.connect(self.addNameClicked)

        # Create the button for adding data names.
        btnAddTCOM = QPushButton('Add TCOM Name', self)
        btnAddTCOM.move(630, 120)
        btnAddTCOM.resize(150, 100)
        btnAddTCOM.clicked.connect(self.addTCOMClicked)

        # Create the button for saving data names.
        btnSaveExit = QPushButton('Save && Exit', self)
        btnSaveExit.move(630, 230)
        btnSaveExit.resize(150, 100)
        btnSaveExit.clicked.connect(self.saveExit)

        # Create the button for canceling changes.
        btnCancelExit = QPushButton('Cancel && Exit', self)
        btnCancelExit.move(630, 340)
        btnCancelExit.resize(150, 100)
        btnCancelExit.clicked.connect(self.cancelExit)

        # Populate the tree via the existing field mappings.
        # Lookup table loaded from .xlsx during initial GUI setup.
        # Make the items editable via double-click.
        for colName in list(fieldMappings):
            dataCol = QTreeWidgetItem([colName])
            self.colTree.addTopLevelItem(dataCol)
            for rawName in fieldMappings[colName].dropna():
                newChild = QTreeWidgetItem([rawName])
                newChild.setFlags(newChild.flags() | QtCore.Qt.ItemIsEditable)
                dataCol.addChild(newChild)
        self.colTree.setCurrentItem(dataCol)

    def addNameClicked(self):
        """Add new tag to a TCOM master data column."""
        # Check if we've selected a TCOM name to add tag to.
        if not self.colTree.currentIndex().parent().isValid():
            text, ok = QInputDialog.getText(self, 'Add Data Name',
                                            'Enter new commission file name for ' 
                                            + self.colTree.currentItem().text(0)
                                            + ':')
            # Check to see if we've entered text.
            if ok and text.strip() != '':
                currentTCOM = self.colTree.currentItem()
                newChild = QTreeWidgetItem([text])
                newChild.setFlags(newChild.flags() | QtCore.Qt.ItemIsEditable)
                currentTCOM.addChild(newChild)

    def addTCOMClicked(self):
        """Add new TCOM master column."""
        text, ok = QInputDialog.getText(self, "Add TCOM Name",
                                        "Enter new TCOM column name:")
        # Check to see if we've entered text.
        if ok and text.strip() != '':
            newTCOM = QTreeWidgetItem([text])
            self.colTree.addTopLevelItem(newTCOM)

    def keyPressEvent(self, event):
        """Wire delete key for expected functionality."""
        # Allow delete key to remove items at all levels.
        if event.key() == QtCore.Qt.Key_Delete:
            root = self.colTree.invisibleRootItem()
            for item in self.colTree.selectedItems():
                (item.parent() or root).removeChild(item)

    def saveExit(self):
        """Save changes to field mappings and close window."""
        global fieldMappings
        fieldMappings = pd.DataFrame()

        # Save tree to application space.
        # Iterate over branches to rebuild field mappings.
        root = self.colTree.invisibleRootItem()
        for colNum in range(root.childCount()):
            newCol = pd.DataFrame(columns=[root.child(colNum).text(0)])
            for childNum in range(root.child(colNum).childCount()):
                newCol = newCol.append(
                        {root.child(colNum).text(0): root.child(colNum).child(childNum).text(0)},
                        ignore_index=True)
            fieldMappings = pd.concat([fieldMappings, newCol], axis=1)

        # Save tree to .xlsx file.
        writer = pd.ExcelWriter('fieldMappings.xlsx', engine='xlsxwriter')
        fieldMappings.to_excel(writer, sheet_name='Lookup', index=False)
        try:
            writer.save()
        except IOError:
            print('Field mappings is open in Excel!')
            print('Please close fieldMappings.xlsx and try again.')
            print('***')
            return

        # Save and exit if no error.
        writer.save()
        print('Changes saved to field mappings.')
        print('---')
        # Close window.
        self.close()

    def cancelExit(self):
        """Close the window without saving changes to field mappings."""
        # Close window. Nothing gets saved.
        print('Mapping changes canceled.')
        print('---')
        self.close()

    def closeEvent(self, event):
        """Close event."""
        # Restore buttons in main GUI on close.
        gui.restoreButtons()


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
