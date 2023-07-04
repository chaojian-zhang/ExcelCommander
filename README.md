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

Reading Routines: 

* Get range
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
* GetNumberFormat range
* GetNumberFormat row, col
* GetCellValues cell, rows, cols
* GetCellValues range
* GetSheet sheetName
* GetTable tableName
* HasNamedRange name
* HasSheet name
* HasTable name

Writing Routines: 

* Align range, option
* Background range, color
* Bold range
* Bold range, toggle
* Border range, weight
* Cell range, value
* Clear range
* ClearFormat range
* Color range, color
* Color row, col, color
* CreateSheet sheetName
* CreateTable range, tableName
* CSV start, filename
* DeleteColumn column
* DeleteColumns columnRange
* DeleteRow row
* DeleteRows rowRange
* DeleteSheet sheetName
* Filter tableOrRange, column, values
* Fit range
* Formula cell, equation
* Formula row, col, equation
* Italic range
* Italic range, toggle
* Merge range
* MoveSheetBefore sheetName, otherSheetName
* NameRange range, name
* NumberFormat range, nameOrFormat
* Outline range
* RenameSheet newName
* RenameSheet originalName, newName
* SetCell cell, value
* SetCell row, col, value
* SetCellName cell, name
* SetCellName row, col, name
* SetCellValues start, csv
* SetFontColor range, color
* SetFontSize range, size
* Style range, name
* Width range, width
* Wrap range
* Wrap range, toggle

State Management Routines: 

* Select range
* GoToSheet sheetName

Macro: 

* Apply range
* Fill range
* Fill from, to
* FillTo range, direction
* InsertRow before
* InsertColumn before
* Paste range
* Save outputFilePath
* Sort range

Programming: 

* Evaluate scriptPath

Utilities: 

* Random range
* Random range, multiplier
* Random range, from, to

Finance: 

* ETL range, outputCell
* ETL range, outputCell, percentage

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

### Color Names

* Transparent
* AliceBlue
* AntiqueWhite
* Aqua
* Aquamarine
* Azure
* Beige
* Bisque
* Black
* BlanchedAlmond
* Blue
* BlueViolet
* Brown
* BurlyWood
* CadetBlue
* Chartreuse
* Chocolate
* Coral
* CornflowerBlue
* Cornsilk
* Crimson
* Cyan
* DarkBlue
* DarkCyan
* DarkGoldenrod
* DarkGray
* DarkGreen
* DarkKhaki
* DarkMagenta
* DarkOliveGreen
* DarkOrange
* DarkOrchid
* DarkRed
* DarkSalmon
* DarkSeaGreen
* DarkSlateBlue
* DarkSlateGray
* DarkTurquoise
* DarkViolet
* DeepPink
* DeepSkyBlue
* DimGray
* DodgerBlue
* Firebrick
* FloralWhite
* ForestGreen
* Fuchsia
* Gainsboro
* GhostWhite
* Gold
* Goldenrod
* Gray
* Green
* GreenYellow
* Honeydew
* HotPink
* IndianRed
* Indigo
* Ivory
* Khaki
* Lavender
* LavenderBlush
* LawnGreen
* LemonChiffon
* LightBlue
* LightCoral
* LightCyan
* LightGoldenrodYellow
* LightGreen
* LightGray
* LightPink
* LightSalmon
* LightSeaGreen
* LightSkyBlue
* LightSlateGray
* LightSteelBlue
* LightYellow
* Lime
* LimeGreen
* Linen
* Magenta
* Maroon
* MediumAquamarine
* MediumBlue
* MediumOrchid
* MediumPurple
* MediumSeaGreen
* MediumSlateBlue
* MediumSpringGreen
* MediumTurquoise
* MediumVioletRed
* MidnightBlue
* MintCream
* MistyRose
* Moccasin
* NavajoWhite
* Navy
* OldLace
* Olive
* OliveDrab
* Orange
* OrangeRed
* Orchid
* PaleGoldenrod
* PaleGreen
* PaleTurquoise
* PaleVioletRed
* PapayaWhip
* PeachPuff
* Peru
* Pink
* Plum
* PowderBlue
* Purple
* Red
* RosyBrown
* RoyalBlue
* SaddleBrown
* Salmon
* SandyBrown
* SeaGreen
* SeaShell
* Sienna
* Silver
* SkyBlue
* SlateBlue
* SlateGray
* Snow
* SpringGreen
* SteelBlue
* Tan
* Teal
* Thistle
* Tomato
* Turquoise
* Violet
* Wheat
* White
* WhiteSmoke
* Yellow
* YellowGreen

### Fill Directions

* Up
* Down
* Left
* Right