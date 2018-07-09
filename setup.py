import sys
import os.path
from cx_Freeze import setup, Executable


"""NOTES:

You will need to copy the 'platforms' folder in the anaconda3 directory
to the directory containing the exe file for the GUI to work.
"""

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os", "numpy"],
                     "includes": ["numpy"]}

# Set environment variables.
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')

# GUI applications require a different base on Windows (the default is for a
# console application).
base = "Win32GUI"

setup(name="Commissions Manager 2.0",
      version="0.1",
      description="",
      options={"build_exe": build_exe_options},
      executables=[Executable("GenerateMasterGUI.py", base=base)])
