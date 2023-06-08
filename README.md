# ExcelCommander

## Setup

(PENDING)

## Architecture

There are four distinct uses

* Repl interactively using either ExcelCommander or ElsxCommander; The ICommander interface guarantees same call signatures.
* Write text-based scripts and execute in either ExcelCommander or ElsxCommander; The ICommander interface guarantees same call signatures.
* Make use of either ExcelCommander.Base, ExcelCommander or ElsxCommander in C# through Pure or Nugets.
* Make use of either ExcelCommander.Base, ExcelCommander or ElsxCommander in Python through PythonNet.

## Usage

### C# Use

* Below snippet is using [Pure](https://github.com/Pure-the-Language/Pure)
* Require `ExcelCommander.exe` folder defined in `PYTHONPATH`
* For regular C# use, add NuGet package as project reference

```C#
Import(ExcelCommander)
using ExcelCommander;
var commander = ExcelCommander.ExcelCommander.Connect(57289);

for(int i = 1; i < 16; i++)
{
	commander.SetCell($"A{i}", i.ToString());
	WriteLine($"A{i}");
}
```

### Python Use

* Require [`pythonnet`](https://pypi.org/project/pythonnet/)
* Require `ExcelCommander.exe` and `excelcommander.py` folder defined in `PYTHONPATH`

```Python
from excelcommander import *
connection = ExcelCommander.Connect(61480)
connection.SetCell("A2", "15")
```

## Supported Commands

* GetCell cell
* GetCell row, col
* GetCellColor cell
* GetCellColor row, col
* GetCellFontWeight cell
* GetCellFontWeight row, col
* GetCellFormula cell
* GetCellFormula row, col
* GetCellName cell
* GetCellName row, col
* GetCellValue cell
* GetCellValue row, col
* GetCellValueFormat cell
* GetCellValueFormat row, col
* GetCellValues cell, rows, cols
* GetCellValues range
* GetCellValues startcell, endcell
* GetCurrentSheet 
* GetSheet sheetName
* GetSheets 
* GetTable tableName
* HasNamedRange name
* HasSheet name
* HasTable name
* Background range, color
* Bold range, weight
* Clear range
* ClearFormat range
* CreateSheet sheetName
* CreateTable range, tableName
* CSV start, filename
* Fit range
* MoveSheetBefore sheetName, otherSheetName
* NameRange range, name
* SetCell cell, value
* SetCell row, col, value
* SetCellName cell, name
* SetCellName row, col, name
* SetCellValues start, csv
* SetColor cell, color
* SetColor row, col, color
* SetEquation cell, equation
* SetEquation row, col, equation
* SetFontColor range, color
* SetFontSize range, size
* SetValueFormat range, format
* GoToSheet sheetName

In C#/Python use, call explicit functions through `ExcelCommander` or `XlsxCommander`, or use `ExecuteCommand()` method.

## Reference

### Alignment Options

* Center
* CenterAcrossSelection
* Distributed
* General: Align according to data type.
* Justify
* Left
* Right

### Border Options

Weights:

* Hairline: thinnest border
* Medium
* Thick
* Thin

### Fill Directions

* Up
* Down
* Left
* Right