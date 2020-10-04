import sys
import os.path
from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import QMainWindow, QPushButton, QApplication, QFileDialog, QTextEdit
from PyQt5.QtCore import pyqtSlot
import LookupSales
import CompileFeedback

LOOKUPS_DIR = 'Z:/Commissions Lookup/'
if not os.path.exists(LOOKUPS_DIR):
    LOOKUPS_DIR = os.getcwd()


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
        sys.stdout = Stream(newText=self.on_update_text)
        # Show welcome message.
        print('Welcome to the TAARCOM Digikey Insights Manager.\n'
              'Version m.10042020\n'
              'Messages and updates will display below.\n'
              '______________________________________________________\n'
              'REMINDER: Did you pull the latest version from GitHub?\n'
              '---')

        # Try finding/loading the supporting files.
        if not os.path.exists(os.path.join(LOOKUPS_DIR, 'rootCustomerMappings.xlsx')):
            print('No Root Customer Mappings found!\n'
                  'Please make sure rootCustomerMappings is in the directory.\n***')

    def on_update_text(self, text):
        """Write console output to text widget."""
        cursor = self.text_box.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.text_box.setTextCursor(cursor)
        self.text_box.ensureCursorVisible()

    def closeEvent(self, event):
        """Shuts down application on close."""
        # Return stdout to defaults.
        sys.stdout = sys.__stdout__
        super().closeEvent(event)

    def initUI(self):
        """Creates UI window on launch."""

        # Button for looking up salespeople in Insight Report.
        self.btn_lookup_sales = QPushButton('Lookup \n Salespeople \n for \n Insight File', self)
        self.btn_lookup_sales.move(650, 400)
        self.btn_lookup_sales.resize(150, 150)
        self.btn_lookup_sales.clicked.connect(self.lookup_sales_clicked)
        self.btn_lookup_sales.setToolTip('Returns a Digikey LI file with salespeople looked '
                                         'up from Account List and rootCustomerMappings.')

        # Button for copying over comments.
        self.btn_compile_feedback = QPushButton('Compile \n Feedback', self)
        self.btn_compile_feedback.move(650, 200)
        self.btn_compile_feedback.resize(150, 150)
        self.btn_compile_feedback.clicked.connect(self.compile_feedback_clicked)
        self.btn_compile_feedback.setToolTip('Combine individual reports with '
                                             'feedback into one file, and '
                                             'append that file to Digikey Insight Master.')

        # Button for clearing selections.
        self.btn_clear_all = QPushButton('Clear \n File Selections', self)
        self.btn_clear_all.move(450, 30)
        self.btn_clear_all.resize(150, 100)
        self.btn_clear_all.clicked.connect(self.clear_all_clicked)

        # Button for selecting new file to lookup salespeople.
        self.btn_open_insight = QPushButton('Select New \n Digikey Insight \n File', self)
        self.btn_open_insight.move(50, 30)
        self.btn_open_insight.resize(150, 100)
        self.btn_open_insight.clicked.connect(self.open_insight_clicked)
        self.btn_open_insight.setToolTip('Select a brand new Digikey LI file.')

        # Button for selecting files to append to master.
        self.btn_open_finished = QPushButton('Select Files \n with Feedback', self)
        self.btn_open_finished.move(250, 30)
        self.btn_open_finished.resize(150, 100)
        self.btn_open_finished.clicked.connect(self.open_finished_clicked)
        self.btn_open_finished.setToolTip('Select a batch of finished files '
                                          'that have feedback from salespeople.')

        # Create the text output widget.
        self.text_box = QTextEdit(self, readOnly=True)
        self.text_box.ensureCursorVisible()
        self.text_box.setLineWrapColumnOrWidth(500)
        self.text_box.setLineWrapMode(QTextEdit.FixedPixelWidth)
        self.text_box.setFixedWidth(550)
        self.text_box.setFixedHeight(400)
        self.text_box.move(50, 150)

        # Set window size and title, then show the window.
        self.setGeometry(300, 300, 900, 600)
        self.setWindowTitle('Digikey Insights Manager 2.0')
        self.show()

    def lookup_sales_clicked(self):
        """Send the LookupSales execution to a worker thread."""
        self.lock_buttons()
        worker = Worker(self.look_sales_execute)
        if self.threadpool.activeThreadCount() == 0:
            self.threadpool.start(worker)

    def compile_feedback_clicked(self):
        """Send the WriteComments execution to a worker thread."""
        self.lock_buttons()
        worker = Worker(self.compile_feedback_execute)
        if self.threadpool.activeThreadCount() == 0:
            self.threadpool.start(worker)

    def clear_all_clicked(self):
        """Clear the filename variables."""
        self.filenames = []
        self.filename = []
        print('All file selections cleared.\n---')
        self.restore_buttons()

    def compile_feedback_execute(self):
        """Runs function for compiling feedback and appending to Master."""
        # Check to make sure we've selected files.
        if self.filenames:
            # Run the GenerateMaster.py file.
            try:
                CompileFeedback.main(self.filenames)
            except Exception as error:
                print('Unexpected Python error:\n'
                      + str(error) + '\nPlease contact your local coder.')
            # Clear files.
            self.filenames = []
        elif not self.filenames:
            print('No finished Insight files selected!\n'
                  'Use the Select Commission Files button to select files.\n---')
        self.restore_buttons()

    def look_sales_execute(self):
        """Runs function for looking up salespeople."""
        # Check to see if we're ready to process.
        map_exists = os.path.exists(os.path.join(LOOKUPS_DIR, 'rootCustomerMappings.xlsx'))
        if self.filename and map_exists:
            # Run the GenerateMaster.py file.
            try:
                LookupSales.main(self.filename)
            except Exception as error:
                print('Unexpected Python error:\n' + str(error) + '\nPlease contact your local coder.')
            # Clear file.
            self.filename = []
        elif not map_exists:
            print('File rootCustomerMappings.xlsx not found!\n'
                  'Please check file location and try again.\n---')
        elif not self.filename:
            print('No Insight file selected!\n'
                  'Use the Select Digikey Insight button to select files.\n---')
        self.restore_buttons()

    def open_insight_clicked(self):
        """Provide filepath for new data to process using LookupSales."""
        # Let us know we're clearing old selections.
        if self.filename:
            print('Selecting new file, old selection cleared...')

        # Grab the filenames to be passed into LookupSales.
        self.filename, _ = QFileDialog.getOpenFileName(self, filter="Excel files (*.xls *.xlsx *.xlsm)")

        # Check if the current master got uploaded as a new file.
        if 'Digikey Insight Master' in self.filename:
            print('Master uploaded as new file.\nTry uploading files again.\n---')
            return

        # Print out the selected filenames.
        if self.filename:
            print('File selected: ' + self.filename + '\n---')
            # Turn off/on the correct buttons.
            self.btn_compile_feedback.setEnabled(False)
            self.btn_lookup_sales.setEnabled(True)

    def open_finished_clicked(self):
        """Provide filepath for new data to process using AppendInsights."""
        # Let us know we're clearing old selections.
        if self.filenames:
            print('Selecting new files, old selections cleared...')

        # Grab the filenames to be passed into AppendInsights.
        self.filenames, _ = QFileDialog.getOpenFileNames(self, filter="Excel files (*.xls *.xlsx *.xlsm)")

        # Check if the current master got uploaded as a new file.
        for name in self.filenames:
            if 'Digikey Insight Master' in name:
                print('Master uploaded as new file.\nTry uploading files again.\n---')
                return

        # Print out the selected filenames.
        if self.filenames:
            print('Files selected:')
            for file in self.filenames:
                print(file)
            print('---')
            # Turn off/on the correct buttons.
            self.btn_compile_feedback.setEnabled(True)
            self.btn_lookup_sales.setEnabled(False)

    def lock_buttons(self):
        self.btn_open_insight.setEnabled(False)
        self.btn_lookup_sales.setEnabled(False)
        self.btn_open_finished.setEnabled(False)
        self.btn_clear_all.setEnabled(False)
        self.btn_compile_feedback.setEnabled(False)

    def restore_buttons(self):
        self.btn_open_insight.setEnabled(True)
        self.btn_lookup_sales.setEnabled(True)
        self.btn_open_finished.setEnabled(True)
        self.btn_clear_all.setEnabled(True)
        self.btn_compile_feedback.setEnabled(True)


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
