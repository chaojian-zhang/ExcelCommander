# ExcelCommander

## Architecture

There are four distinct uses

* Repl interactively using either ExcelCommander or ElsxCommander; The ICommander interface guarantees same call signatures.
* Write text-based scripts and execute in either ExcelCommander or ElsxCommander; The ICommander interface guarantees same call signatures.
* Make use of either ExcelCommander.Base, ExcelCommander or ElsxCommander in C# through Pure or Nugets.
* Make use of either ExcelCommander.Base, ExcelCommander or ElsxCommander in Python through PythonNet.

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