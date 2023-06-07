# Require path to the folder containing "ExcelCommander.exe" exist in PYTHONPATH (not PATH)
# PYTHONPATH should also contain folder of this file

from pythonnet import load
load("coreclr")

import clr
clr.AddReference("ExcelCommander")

from ExcelCommander import ExcelCommander