import sys
import pandas as pd
import os.path
import logging
import traceback
from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import QMainWindow, QPushButton, QApplication, QFileDialog, QWidget, QGridLayout, \
    QPlainTextEdit, QGroupBox, QVBoxLayout, QListWidget, QListWidgetItem, QLabel
from PyQt5.QtCore import pyqtSlot
import GenerateMaster
import MergeFixedEntries
import SalesReportGenerator
import MergeByUuid
import CommTools
from GenerateMasterUtils import DIRECTORIES

VERSION = 'Master v3.0.040424'
LOOKUPS_DIR = DIRECTORIES.get('COMM_LOOKUPS_DIR')


class GenerateMasterGUI(QMainWindow):
    """Main application window."""
    def __init__(self):
        super().__init__()
        self.main_widget = QWidget(self)
        self.setCentralWidget(self.main_widget)
        main_layout = QGridLayout()
        self.main_widget.setLayout(main_layout)

        # Create the buttons.
        self.button_generate_master = QPushButton('Process Raw Files', self)
        self.button_open_files = QPushButton('Select Raw Commission File(s)', self)
        self.button_load_run_com = QPushButton('Select Running Commissions File', self)
        self.button_fix_entries = QPushButton('Copy Fixed ENF to Running Commissions', self)
        self.button_generate_reports = QPushButton('Migrate and Generate Sales Reports', self)
        self.button_clear_selections = QPushButton('Clear Selections', self)
        self.button_clear_master = QPushButton('Clear Selection', self)
        self.button_update_lookup = QPushButton('Update Lookup Master', self)
        self.button_update_master = QPushButton('Update Commission Master', self)

        self.log_text_box = QTextEditLogger(self)
        self.log_text_box.setFormatter(logging.Formatter(fmt='%(asctime)s (%(levelname)s): %(message)s',
                                                         datefmt="%Y-%m-%d %H:%M:%S"))

        # Create a nice labeled box for the logs and initialize the logger.
        self.log_groupbox = QGroupBox('Log')
        log_layout = QVBoxLayout()
        log_layout.addWidget(self.log_text_box.text_widget)
        self.log_groupbox.setLayout(log_layout)
        logging.getLogger().addHandler(self.log_text_box)
        logging.getLogger().setLevel(logging.DEBUG)

        # Create box for uploading and displaying raw data files.
        self.raw_files_groupbox = QGroupBox('Raw Files')
        self.raw_files_list = QListWidget()
        files_title = QLabel()
        files_title.setText('Current loaded files:')
        raw_files_layout = QVBoxLayout()
        raw_files_layout.addWidget(self.button_open_files)
        raw_files_layout.addWidget(self.button_clear_selections)
        raw_files_layout.addWidget(files_title)
        raw_files_layout.addWidget(self.raw_files_list)
        self.raw_files_groupbox.setLayout(raw_files_layout)

        # Create box for uploading and displaying running commissions file.
        self.runcom_file_groupbox = QGroupBox('Running Commissions File')
        self.runcom_file_list = QListWidget()
        runcom_file_layout = QVBoxLayout()
        runcom_file_layout.addWidget(self.button_load_run_com)
        runcom_file_layout.addWidget(self.button_clear_master)
        runcom_file_layout.addWidget(self.runcom_file_list)
        self.runcom_file_groupbox.setLayout(runcom_file_layout)

        # Create box for program buttons.
        self.programs_groupbox = QGroupBox('Programs')
        programs_layout = QGridLayout()
        programs_layout.addWidget(self.button_generate_master, 0, 0, 1, 2)
        programs_layout.addWidget(self.button_update_lookup, 1, 0)
        programs_layout.addWidget(self.button_update_master, 1, 1)
        programs_layout.addWidget(self.button_fix_entries, 2, 0)
        programs_layout.addWidget(self.button_generate_reports, 2, 1)
        self.programs_groupbox.setLayout(programs_layout)

        # Add GUI elements to the top level grid layout.
        main_layout.addWidget(self.raw_files_groupbox, 0, 0, 1, 2)
        main_layout.addWidget(self.runcom_file_groupbox, 1, 0, 1, 2)
        main_layout.addWidget(self.programs_groupbox, 2, 0, 1, 2)
        main_layout.addWidget(self.log_groupbox, 0, 3, 5, 5)

        # Set the grid geometry.
        main_layout.setRowStretch(0, 2)
        main_layout.setRowStretch(1, 1)
        main_layout.setRowStretch(2, 1)

        # Set window size and title, then show the window.
        self.setGeometry(100, 100, 2000, 1200)
        self.setWindowTitle(f'Commissions Manager {VERSION}')
        self.show()

        # Set working directory as current file directory.
        os.chdir(os.path.dirname(os.path.realpath(__file__)))
        # Initialize the threadpool for handling worker jobs.
        self.threadpool = QtCore.QThreadPool()

        # Create the widgets and layout and launch the main window.
        self.connect_buttons()

        self.filenames = []
        self.master = []

        logging.info(f'Welcome to the TAARCOM Commissions Manager!\nVersion: {VERSION}\n'
                     'Logs will display here.\n'
                     '_______________________________________________________________\n'
                     'REMINDER: Did you check for updates on GitHub?\n'
                     'REMINDER: If new code was pulled or the branch changed, please close and relaunch the program.\n')

        if not os.path.exists('Z:/'):
            logging.info('No connection to Z:/ drive established! Working locally.')

        # Try loading/finding the supporting files.
        if os.path.exists(LOOKUPS_DIR):
            if os.path.exists(os.path.join(LOOKUPS_DIR, 'fieldMappings.xlsx')):
                self.fieldMappings = pd.read_excel(os.path.join(LOOKUPS_DIR, 'fieldMappings.xlsx'),
                                                   index_col=False)
            else:
                logging.warning(f'No field mappings found! '
                                f'Please make sure fieldMappings.xlsx is located at {LOOKUPS_DIR}')

            if not os.path.exists(os.path.join(LOOKUPS_DIR, 'Lookup Master - Current.xlsx')):
                logging.warning('No Lookup Master found! '
                                f'Please make sure Lookup Master is located at {LOOKUPS_DIR}')

            if not os.path.exists(os.path.join(LOOKUPS_DIR, 'distributorLookup.xlsx')):
                logging.warning('No distributor lookup found!'
                                f'Please make sure distributorLookup.xlsx is located at {LOOKUPS_DIR}')
        else:
            logging.warning(f'Could not connect to {LOOKUPS_DIR} Please make sure you '
                            'are connected to the TAARCOM server, then relaunch the program.')

    def closeEvent(self, event):
        """Shuts down application on close."""
        super().closeEvent(event)

    def connect_buttons(self):
        """Creates UI window on launch."""
        # Button for generating the master list.
        self.button_generate_master.clicked.connect(self.generate_master_clicked)
        self.button_generate_master.setToolTip('Process selected raw data files and append them to the selected '
                                               'Running Commissions.\nIf no Running Commissions is selected, '
                                               'starts a new one.')

        # Button for selecting files to compile into master list.
        self.button_open_files.clicked.connect(self.open_files_clicked)
        self.button_open_files.setToolTip('Open explorer tool for selecting raw commission file(s) to process.')

        # Button for selecting a current master to append to.
        self.button_load_run_com.clicked.connect(self.upload_master_clicked)
        self.button_load_run_com.setToolTip('Open explorer tool for selecting a Running Commissions to use.')

        # Button for writing fixed entries.
        self.button_fix_entries.clicked.connect(self.fix_entries_clicked)
        self.button_fix_entries.setToolTip('Migrate finished lines in the Entries Need Fixing file over to the '
                                           'associated Running Commissions.\nA Running Commissions needs to be '
                                           'selected and will be matched to the Entries Needs Fixing by the '
                                           'date at the end of the filenames.')

        # Button for generating sales reports.
        self.button_generate_reports.clicked.connect(self.generate_reports_clicked)
        self.button_generate_reports.setToolTip('Generate commission and revenue reports, then migrate the Running '
                                                'Commissions data over to the Commissions Master.\n'
                                                'If no Running Commissions is provided, will run reports on most '
                                                'recent quarter of data in the Commissions Master.')

        # Button for clearing filename and master choices.
        self.button_clear_selections.clicked.connect(self.clear_files_clicked)
        self.button_clear_selections.setToolTip('Clear all selected raw files from the workspace.')

        # Button for clearing filename and master choices.
        self.button_clear_master.clicked.connect(self.clear_master_clicked)
        self.button_clear_master.setToolTip('Clear selected RC file from the workspace.')

        # Button for updating Lookup Master.
        self.button_update_lookup.clicked.connect(self.update_lookup_clicked)
        self.button_update_lookup.setToolTip('Update the Lookup Master using a Running Commissions file.')

        # Button for updating data in Commission Master.
        self.button_update_master.clicked.connect(self.update_master_clicked)
        self.button_update_master.setToolTip('Update the Commission Master using a Running Commissions file, '
                                             'matching entries by their Unique ID.')

    def generate_master_clicked(self):
        """Send the GenerateMaster execution to a worker thread."""
        self.lock_buttons()
        worker = Worker(self.execute_generate_master)
        if self.threadpool.activeThreadCount() == 0:
            self.threadpool.start(worker)
        else:
            logging.warning('Worker thread busy, aborting program.')

    def generate_reports_clicked(self):
        """Send the SalesReportGenerator execution to a worker thread."""
        self.lock_buttons()
        worker = Worker(self.execute_generate_reports)
        if self.threadpool.activeThreadCount() == 0:
            self.threadpool.start(worker)
        else:
            logging.warning('Worker thread busy, aborting program.')

    def fix_entries_clicked(self):
        """Send the MergeFixedEntries execution to a worker thread."""
        self.lock_buttons()
        worker = Worker(self.execute_fix_entries)
        if self.threadpool.activeThreadCount() == 0:
            self.threadpool.start(worker)
        else:
            logging.warning('Worker thread busy, aborting program.')

    def update_lookup_clicked(self):
        """Send the UpdateLookups execution to a worker thread."""
        self.lock_buttons()
        worker = Worker(self.execute_update_lookups)
        if self.threadpool.activeThreadCount() == 0:
            self.threadpool.start(worker)
        else:
            logging.warning('Worker thread busy, aborting program.')

    def update_master_clicked(self):
        """Send the UpdateLookups execution to a worker thread."""
        self.lock_buttons()
        worker = Worker(self.execute_update_master)
        if self.threadpool.activeThreadCount() == 0:
            self.threadpool.start(worker)
        else:
            logging.warning('Worker thread busy, aborting program.')

    def clear_files_clicked(self):
        """Clear the filenames and master variables."""
        self.filenames = []
        self.raw_files_list.clear()

    def clear_master_clicked(self):
        """Clear the filenames and master variables."""
        self.master = []
        self.runcom_file_list.clear()

    def execute_generate_master(self):
        """Runs function for processing new files to master."""
        # Check to see if we're ready to process.
        mappings_exist = os.path.exists(os.path.join(LOOKUPS_DIR, 'fieldMappings.xlsx'))
        if self.filenames and mappings_exist:
            # Run the GenerateMaster.py file.
            try:
                success = GenerateMaster.main(self.filenames, self.master, self.fieldMappings)
                if success:
                    logging.info('+Program Generate Master Complete!+')
                    self.filenames = []
                    self.raw_files_list.clear()
            except Exception:
                logging.error(f'Unexpected Python error: {traceback.format_exc(0)}\n*Program Terminated*')
        elif not mappings_exist:
            logging.warning('File fieldMappings.xlsx not found! Please check file location and try again.')
        elif not self.filenames:
            logging.warning('No commission files selected! Use the Select Commission Files button to select files.')
        self.restore_buttons()

    def execute_generate_reports(self):
        """Runs function for generating salesperson reports."""
        # Run the SalesReportGenerator.py file.
        try:
            SalesReportGenerator.main(self.master)
        except Exception:
            logging.error(f'Unexpected Python error: {traceback.format_exc(0)}')
        # Turn buttons back on.
        self.restore_buttons()

    def execute_fix_entries(self):
        """Copy over fixed entries to Master."""
        if self.master:
            # Run the GenerateMaster.py file.
            try:
                MergeFixedEntries.main(self.master)
            except Exception as error:
                logging.error(f'Unexpected Python error: {traceback.format_exc(0)}')
        else:
            logging.warning('Please upload the current Running Commissions file and try again.')
        self.restore_buttons()

    def execute_update_lookups(self):
        """Update the Lookup Master using a Running Commissions."""
        if self.master:
            # Run the GenerateMaster.py file.
            try:
                CommTools.extract_lookups(self.master)
            except Exception:
                logging.error(f'Unexpected Python error: {traceback.format_exc(0)}')
        else:
            logging.warning('Please upload a Running Commissions file and try again.')
        self.restore_buttons()

    def execute_update_master(self):
        """Run the update Master Commissions via Unique ID function."""
        if self.master:
            # Run the GenerateMaster.py file.
            try:
                MergeByUuid.main(self.master)
            except Exception:
                logging.error(f'Unexpected Python error: {traceback.format_exc(0)}')
        else:
            logging.warning('Please upload the target Running Commissions file and try again.')
        self.restore_buttons()

    def upload_master_clicked(self):
        """Upload an existing Running Commissions."""
        # Grab an existing Running Commissions to append to.
        self.master, _ = QFileDialog.getOpenFileName(self, filter='Excel files (*.xls *.xlsx *.xlsm)')
        if self.master:
            self.runcom_file_list.clear()
            self.runcom_file_list.addItem(QListWidgetItem(self.master))
            if 'Running Commissions' not in self.master:
                logging.warning('Caution! The file uploaded as a Running Commissions does not appear to be correct.')

    def open_files_clicked(self):
        """Provide filepaths for new data to process using GenerateMaster."""
        # Grab the filenames to be passed into GenerateMaster.
        self.filenames, _ = QFileDialog.getOpenFileNames(self, filter='Excel files (*.xls *.xlsx *.xlsm)')
        # Check if the current master got uploaded as a new file.
        for name in self.filenames:
            if 'Running Commissions' in name:
                logging.warning('Detected RC uploaded as raw file. Try uploading files again.')
                return
        # Print out the selected filenames.
        if self.filenames:
            self.raw_files_list.clear()
            for file in self.filenames:
                self.raw_files_list.addItem(QListWidgetItem(file))

    def lock_buttons(self):
        self.button_generate_master.setEnabled(False)
        self.button_open_files.setEnabled(False)
        self.button_load_run_com.setEnabled(False)
        self.button_clear_selections.setEnabled(False)
        self.button_clear_master.setEnabled(False)
        self.button_fix_entries.setEnabled(False)
        self.button_generate_reports.setEnabled(False)
        self.button_update_lookup.setEnabled(False)
        self.button_update_master.setEnabled(False)

    def restore_buttons(self):
        self.button_generate_master.setEnabled(True)
        self.button_open_files.setEnabled(True)
        self.button_load_run_com.setEnabled(True)
        self.button_clear_selections.setEnabled(True)
        self.button_clear_master.setEnabled(True)
        self.button_fix_entries.setEnabled(True)
        self.button_generate_reports.setEnabled(True)
        self.button_update_lookup.setEnabled(True)
        self.button_update_master.setEnabled(True)


class QTextEditLogger(logging.Handler, QtCore.QObject):
    append_plain_text = QtCore.pyqtSignal(str)

    def __init__(self, parent):
        super().__init__()
        QtCore.QObject.__init__(self)
        self.text_widget = QPlainTextEdit(parent)
        self.text_widget.setReadOnly(True)
        self.append_plain_text.connect(self.text_widget.appendPlainText)

    def emit(self, record):
        msg = self.format(record)
        self.append_plain_text.emit(msg)


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
    gui = GenerateMasterGUI()
    sys.exit(app.exec_())
