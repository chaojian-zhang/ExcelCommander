# Generate Command Doc

This script parses ICommandHandler.cs file to generate command documentation in MD format.

Usage: `pure GenerateCommandDoc <filepath>`

```C#
using System.Text.RegularExpressions;

string scriptFolder = Directory.GetCurrentDirectory();
string solutionFolder = Path.GetDirectoryName(scriptFolder);
string interfaceClassFilePath = Path.Join(solutionFolder, "ExcelCommander.Base", "ICommander.cs");

string text = File.ReadAllText(interfaceClassFilePath);

MatchCollection regions = Regex.Matches(text, @"#region (.+?)\n(.+?)#endregion", RegexOptions.Singleline);
foreach(Match region in regions)
{
	WriteLine(region.Groups[1].Value.TrimEnd() + ": ");
	WriteLine();
	
	string body = region.Groups[2].Value;
	MatchCollection commands = Regex.Matches(body, @"CommandData (.+?)\((string (.+?))+\)", RegexOptions.None);
	foreach(Match command in commands)
		WriteLine("* " + command.Groups[1].Value + " " + string.Join(", ", command.Groups[3].Captures.Select(c => c.Value.Trim().TrimEnd(','))));
	WriteLine();
}
```

```Cache
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
* GetCellValueFormat cell
* GetCellValueFormat row, col
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
* Fit range
* Formula cell, equation
* Formula row, col, equation
* Italic range
* Italic range, toggle
* Merge range
* MoveSheetBefore sheetName, otherSheetName
* NameRange range, name
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
* SetValueFormat range, format
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
```