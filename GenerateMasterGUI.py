import sys
import pandas as pd
import os.path
from PyQt5.QtWidgets import QMainWindow, QPushButton, QApplication, \
                            QFileDialog, QTextEdit, QTreeWidget, \
                            QTreeWidgetItem, QInputDialog
from PyQt5 import QtCore, QtGui
import GenerateMaster


class Stream(QtCore.QObject):
    """Redirects console output to text widget."""
    newText = QtCore.pyqtSignal(str)

    def write(self, text):
        self.newText.emit(str(text))


class GenMast(QMainWindow):
    """Main application window."""
    def __init__(self):
        super().__init__()

        self.initUI()
        self.filenames = []
        self.master = []

        # Create a global varaiable for the lookup table.
        # We can now edit it in the ColumnEdit class, or just leave it alone.
        global lookupTable
        # Upload lookup table, if found.
        if os.path.exists('lookupTable.csv'):
            lookupTable = pd.read_csv('lookupTable.csv', index_col=False)

        # Custom output stream.
        sys.stdout = Stream(newText=self.onUpdateText)

    def onUpdateText(self, text):
        """Write console output to text widget."""
        # Print console output to text box widget.
        cursor = self.process.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.process.setTextCursor(cursor)
        self.process.ensureCursorVisible()

    def closeEvent(self, event):
        """Shuts down application on close."""
        # Return stdout to defaults.
        sys.stdout = sys.__stdout__
        super().closeEvent(event)

    def initUI(self):
        """Creates UI window on launch."""
        # Button for generating the master list.
        btnGenMast = QPushButton('Process Files \n to Master', self)
        btnGenMast.move(450, 100)
        btnGenMast.resize(100, 100)
        btnGenMast.clicked.connect(self.genMastClicked)

        # Button for selecting files to compile into master list.
        btnOpenFiles = QPushButton('Add New Files', self)
        btnOpenFiles.move(30, 50)
        btnOpenFiles.clicked.connect(self.openFilesClicked)

        # Button for selecting files to compile into master list.
        btnUploadMast = QPushButton('Upload Master', self)
        btnUploadMast.move(150, 50)
        btnUploadMast.clicked.connect(self.uploadMastClicked)

        # Button for editing variable names
        btnEditColumns = QPushButton('Edit Column Tags', self)
        btnEditColumns.move(270, 50)
        btnEditColumns.clicked.connect(self.editColumnsClicked)

        # Create the text output widget.
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

    def editColumnsClicked(self):
        """Opens new window for editing lookup table."""
        # Open new window with data tree and editing processes.
        if os.path.exists('lookupTable.csv'):
            self.columnsWindow = ColumnEdit()
            self.columnsWindow.show()
        else:
            print('No lookup table file found!')
            print('Please make sure lookupTable.csv is in the directory.')
            print('***')

    def genMastClicked(self):
        """Runs function for processing new files to master."""
        # Check to see if we've selected files to process.
        if self.filenames:
            # Run the GenerateMaster.py file.
            GenerateMaster.main(self.filenames, self.master, lookupTable)
        else:
            print('No new files selected!')
            print('Use the Open Files button to select files.')
            print('---')

    def uploadMastClicked(self):
        """Upload an existing master list."""
        # Grab an existing master list to append to.
        self.master, _ = QFileDialog.getOpenFileName(
                self, filter="Excel files (*.xls *.xlsx *.xlsm)")
        if self.master:
            print('Master list provided:')
            print(self.master)
            print('---')
            if 'CurrentMaster' not in self.master:
                print('Caution!')
                print('The file uploaded as master \
                      does not appear to be a master list.')
                print('---')

    def openFilesClicked(self):
        """Provide filepaths for new data to process using GenerateMaster."""
        # Grab the filenames to be passed into GenerateMaster.py
        self.filenames, _ = QFileDialog.getOpenFileNames(
                self, filter="Excel files (*.xls *.xlsx *.xlsm)")

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


class ColumnEdit(QMainWindow):
    """Window for editing lookup table contents."""
    def __init__(self, parent=None):
        """Create UI for window on launch."""
        super().__init__()

        # Set window size and title.
        self.setGeometry(200, 200, 600, 300)
        self.setWindowTitle('Column Name List')

        # Create the tree widget with column names.
        self.colTree = QTreeWidget(self)
        self.colTree.resize(500, 200)
        self.colTree.setColumnCount(1)
        self.colTree.setHeaderLabels(['TCOM Column Names'])

        # Create the button for adding data names.
        btnAddName = QPushButton('Add Lookup Name', self)
        btnAddName.move(10, 220)
        btnAddName.clicked.connect(self.addNameClicked)

        # Create the button for adding data names.
        btnAddTCOM = QPushButton('Add TCOM Name', self)
        btnAddTCOM.move(130, 220)
        btnAddTCOM.clicked.connect(self.addTCOMClicked)

        # Create the button for saving data names.
        btnSaveExit = QPushButton('Save && Exit', self)
        btnSaveExit.move(470, 260)
        btnSaveExit.clicked.connect(self.saveExit)

        # Create the button for canceling changes.
        btnCancelExit = QPushButton('Cancel', self)
        btnCancelExit.move(350, 260)
        btnCancelExit.clicked.connect(self.cancelExit)

        # Populate the tree via the existing lookup table.
        # Lookup table loaded from .csv during initial GUI setup.
        # Make the items editable via double-click.
        for colName in list(lookupTable):
            dataCol = QTreeWidgetItem([colName])
            self.colTree.addTopLevelItem(dataCol)
            for rawName in lookupTable[colName].dropna():
                newChild = QTreeWidgetItem([rawName])
                newChild.setFlags(newChild.flags() | QtCore.Qt.ItemIsEditable)
                dataCol.addChild(newChild)
        self.colTree.setCurrentItem(dataCol)

    def addNameClicked(self):
        """Add new tag to a TCOM master data column."""
        # Check if we've selected a TCOM name to add tag to.
        if not self.colTree.currentIndex().parent().isValid():
            text, ok = QInputDialog.getText(self, "Add Data Name",
                                            "Enter new data name:")
            # Check to see if we've entered text.
            if ok and text != '':
                currentTCOM = self.colTree.currentItem()
                newChild = QTreeWidgetItem([text])
                newChild.setFlags(newChild.flags() | QtCore.Qt.ItemIsEditable)
                currentTCOM.addChild(newChild)

    def addTCOMClicked(self):
        """Add new TCOM master column."""
        text, ok = QInputDialog.getText(self, "Add TCOM Name",
                                        "Enter new TCOM name:")
        # Check to see if we've entered text.
        if ok and text != '':
            newTCOM = QTreeWidgetItem([text])
            self.colTree.addTopLevelItem(newTCOM)

    # Allow delete key to remove items at all levels.
    def keyPressEvent(self, event):
        """Wire delete key for expected functionality."""
        if event.key() == QtCore.Qt.Key_Delete:
            root = self.colTree.invisibleRootItem()
            for item in self.colTree.selectedItems():
                (item.parent() or root).removeChild(item)

    def saveExit(self):
        """Save changes to lookup table and close window."""
        global lookupTable
        lookupTable = pd.DataFrame()

        # Save tree to application space.
        # Iterate over branches to rebuild lookup table.
        root = self.colTree.invisibleRootItem()
        for colNum in range(root.childCount()):
            newCol = pd.DataFrame(columns=[root.child(colNum).text(0)])
            for childNum in range(root.child(colNum).childCount()):
                newCol = newCol.append(
                        {root.child(colNum).text(0): root.child(colNum).child(childNum).text(0)},
                        ignore_index=True)
            lookupTable = pd.concat([lookupTable, newCol], axis=1)

        # Save tree to .csv file.
        lookupTable.fillna('').to_csv('lookupTable.csv', index=False)
        print('Lookup table changes saved.')
        print('---')

        # Close window.
        self.close()

    def cancelExit(self):
        """Close the window without saving changes to lookup table."""
        # Close window. Nothing gets saved.
        self.close()


if __name__ == '__main__':
    # Run the application.
    app = QApplication(sys.argv)
    app.aboutToQuit.connect(app.deleteLater)
    gui = GenMast()
    sys.exit(app.exec_())
