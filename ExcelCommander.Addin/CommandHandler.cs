using ExcelCommander.Base.Serialization;
using ExcelCommander.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Text;

namespace ExcelCommander.Addin
{
    internal class CommandHandler: ICommander
    {
        #region Properties
        private MethodInfo[] _CommandMethods;
        private MethodInfo[] CommandMethods
        {
            get
            {
                if (_CommandMethods == null)
                    _CommandMethods = GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly).ToArray();
                return _CommandMethods;
            }
        }
        #endregion

        #region Preset Replies
        private CommandData Ok() => new CommandData()
        {
            CommandType = "Ok",
            Contents = string.Empty
        };
        private CommandData Error(string message) => new CommandData()
        {
            CommandType = "Error",
            Contents = message ?? string.Empty
        };
        private CommandData Value(bool value) => new CommandData()
        {
            CommandType = "Value bool",
            Contents = value.ToString()
        };
        private CommandData Value(int value) => new CommandData()
        {
            CommandType = "Value int",
            Contents = value.ToString()
        };
        private CommandData Value(double value) => new CommandData()
        {
            CommandType = "Value double",
            Contents = value.ToString()
        };
        private CommandData Value(char value) => new CommandData()
        {
            CommandType = "Value char",
            Contents = value.ToString()
        };
        private CommandData Value(string value) => new CommandData()
        {
            CommandType = "Value string",
            Contents = value
        };
        private CommandData Values(bool[] values) => new CommandData()
        {
            CommandType = "Value bool[]",
            Contents = string.Join(Environment.NewLine, values)
        };
        private CommandData Values(int[] values) => new CommandData()
        {
            CommandType = "Value int[]",
            Contents = string.Join(Environment.NewLine, values)
        };
        private CommandData Values(double[] values) => new CommandData()
        {
            CommandType = "Value double[]",
            Contents = string.Join(Environment.NewLine, values)
        };
        private CommandData Values(char[] values) => new CommandData()
        {
            CommandType = "Value char[]",
            Contents = string.Join(Environment.NewLine, values)
        };
        private CommandData Values(string[] values) => new CommandData()
        {
            CommandType = "Value string[]",
            Contents = string.Join(Environment.NewLine, values)
        };
        #endregion

        #region Entry Point
        internal CommandData Handle(CommandData data)
        {
            if (data.CommandType.Split(' ').First() == nameof(SetCellValues))
                return SetCellValues(data.CommandType.Split(' ').Last(), data.Contents);
            else
                return HandleGeneralCommands();

            CommandData HandleGeneralCommands()
            {
                string[] parameters = data.Contents.SplitParameters(true);

                var methods = CommandMethods;
                var match = methods.FirstOrDefault(m =>
                    m.Name == parameters[0]
                    && m.GetParameters().Length == parameters.Length - 1
                    && m.ReturnType == typeof(CommandData));
                if (match != null)
                    return (CommandData)match.Invoke(this, parameters.Skip(1).OfType<string>().ToArray());  // Remar-cz: Notice we are not trimming `"` right here and instead require all specific handling routines to do it explicitly because sometimes we might want the additional semantics to denote something like "15" as a text instead of pure value
                else return null;
            }
        }
        #endregion

        #region Reading Routines
        public CommandData Get(string range)
        {
            try
            {
                return Value(ToCSV(Application.Range[range].Value2));
            }
            catch (Exception e)
            {
                return Error(e.Message);
            }
        }
        public CommandData GetCell(string cell)
        {
            try
            {
                if (TryGetRowCol(cell, out int row, out int col))
                    return Value(ActiveWorksheet.Cells[row, col].Value.ToString());
            }
            catch (Exception e)
            {
                return Error(e.Message);
            }
            
            return Ok();
        }
        public CommandData GetCell(string row, string col)
        {
            try
            {
                return Value(ActiveWorksheet.Cells[row, col].Value.ToString());
            }
            catch (Exception e)
            {
                return Error(e.Message);
            }
        }
        public CommandData GetCellColor(string cell)
        {
            try
            {
                return Value(ActiveWorksheet.Range[cell].Interior.Color.ToString());
            }
            catch (Exception e)
            {
                return Error(e.Message);
            }
        }
        public CommandData GetCellColor(string row, string col)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCellFontWeight(string cell)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCellFontWeight(string row, string col)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCellFormula(string cell)
        {
            try
            {
                return Value(ActiveWorksheet.Range[cell].Formula);
            }
            catch (Exception e)
            {
                return Error(e.Message);
            }
        }
        public CommandData GetCellFormula(string row, string col)
        {
            try
            {
                return Value(ActiveWorksheet.Cells[int.Parse(row), int.Parse(col)].Formula);
            }
            catch (Exception e)
            {
                return Error(e.Message);
            }
        }
        public CommandData GetCellName(string range)
        {
            try
            {
                return Value(ActiveWorksheet.Range[range].Name);
            }
            catch (Exception e)
            {
                return Error(e.Message);
            }
        }
        public CommandData GetCellName(string row, string col)
        {
            try
            {
                return Value(ActiveWorksheet.Range[int.Parse(row), int.Parse(col)].Name);
            }
            catch (Exception e)
            {
                return Error(e.Message);
            }
        }
        public CommandData GetCellValue(string cell)
        {
            try
            {
                return Value(ActiveWorksheet.Cells[cell].Value2.ToString());
            }
            catch (Exception e)
            {
                return Error(e.Message);
            }
        }
        public CommandData GetCellValue(string row, string col)
        {
            try
            {
                return Value(ActiveWorksheet.Cells[int.Parse(row), int.Parse(col)].Value2.ToString());
            }
            catch (Exception e)
            {
                return Error(e.Message);
            }
        }
        public CommandData GetCellValues(string cell, string rows, string cols)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCellValues(string range)
        {
            try
            {
                return Value(ToCSV(ActiveWorksheet.Cells[range].Value2));
            }
            catch (Exception e)
            {
                return Error(e.Message);
            }
        }
        public CommandData GetCurrentSheet()
        {
            throw new NotImplementedException();
        }
        public CommandData GetNumberFormat(string cell)
        {
            try
            {
                return Value(Application.Range[cell].NumberFormat);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData GetNumberFormat(string row, string col)
        {
            try
            {
                return Value(Application.Cells[int.Parse(row), int.Parse(col)].NumberFormat);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData GetTable(string tableName)
        {
            throw new NotImplementedException();
        }
        public CommandData GetSheet(string sheetName)
        {
            throw new NotImplementedException();
        }
        public CommandData GetSheets()
        {
            try
            {
                string[] names = Enumerable.Range(1, Application.ActiveWorkbook.Worksheets.Count)   // Remark-cz: Excel index starts at 1
                    .Select(i => Application.ActiveWorkbook.Worksheets[i].Name as string)
                    .ToArray();
                return Values(names);
            }
            catch (Exception e)
            {
                return Error(e.Message);
            }
        }
        public CommandData HasNamedRange(string name)
        {
            throw new NotImplementedException();
        }
        public CommandData HasSheet(string name)
        {
            throw new NotImplementedException();
        }
        public CommandData HasTable(string name)
        {
            return Value(HasWorkSheet(name).ToString());
        }
        #endregion

        #region Writing Routines
        public CommandData Align(string range, string option)
        {
            try
            {
                Application.Range[range].HorizontalAlignment = (XlHAlign)Enum.Parse(typeof(XlHAlign), $"xlHAlign{option}"); // Remark-cz: E.g. XlHAlign.xlHAlignLeft
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Background(string range, string color)
        {
            try
            {
                Application.Range[range].Interior.Color = ParseColor(color);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Bold(string range)
        {
            try
            {
                Application.Range[range].Font.Bold = !Application.Range[range].Font.Bold;
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Bold(string range, string toggle)
        {
            try
            {
                Application.Range[range].Font.Bold = bool.Parse(toggle);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Border(string range, string weight)
        {
            try
            {
                var w = (XlBorderWeight)Enum.Parse(typeof(XlBorderWeight), $"xl{weight}");
                ActiveWorksheet.Range[range].Borders.Weight = w;
                ActiveWorksheet.Range[range].Borders.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Cell(string range, string value)
        {
            try
            {
                Application.Range[range].Value = ParseValue(value);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Clear(string range)
        {
            try
            {
                Application.Range[range].Clear();
            }
            catch (Exception) { }
            return null;
        }
        public CommandData ClearAll()
        {
            try
            {
                ActiveWorksheet.UsedRange.Clear();
            }
            catch (Exception) { }
            return null;
        }
        public CommandData ClearFormat(string range)
        {
            try
            {
                Application.Range[range].ClearFormats();
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Color(string range, string color)
        {
            try
            {
                Application.Range[range].Font.Color = ParseColor(color);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Color(string row, string col, string color)
        {
            try
            {
                ActiveWorksheet.Cells[int.Parse(row), int.Parse(col)].Font.Color = ParseColor(color);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData CSV(string start, string filename)
        {
            try
            {
                return SetCellValues(start, System.IO.File.ReadAllText(ParseString(filename)));
            }
            catch (Exception)
            {
                return null;
            }
        }
        public CommandData CreateTable(string range, string tableName)
        {
            try
            {
                ActiveWorksheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, Application.get_Range(range), null, XlYesNoGuess.xlYes, null, "TableStyleMedium3").Name = ParseString(tableName);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData CreateSheet(string sheetName)
        {
            TryCreateWorksheet(ParseString(sheetName));
            return null;
        }
        public CommandData DeleteColumn(string column)
        {
            try
            {
                ActiveWorksheet.Columns[ParseValue(column)].Delete();
            }
            catch (Exception){}
            return null;
        }
        public CommandData DeleteColumns(string columnRange)
        {
            try
            {
                if (columnRange.Contains(":"))
                    ActiveWorksheet.Columns[columnRange].Delete();
                else if (columnRange.Contains(","))
                    foreach (var column in columnRange.Split(','))
                        ActiveWorksheet.Columns[ParseValue(column)].Delete();
            }
            catch (Exception) { }
            return null;
        }
        public CommandData DeleteRow(string row)
        {
            try
            {
                ActiveWorksheet.Rows[ParseValue(row)].Delete();
            }
            catch (Exception) { }
            return null;
        }
        public CommandData DeleteRows(string rowRange)
        {
            try
            {
                if (rowRange.Contains(":"))
                    ActiveWorksheet.Rows[rowRange].Delete();
                else if (rowRange.Contains(","))
                    foreach (var row in rowRange.Split(','))
                        ActiveWorksheet.Rows[ParseValue(row)].Delete();
            }
            catch (Exception) { }
            return null;
        }
        public CommandData DeleteSheet(string sheetName)
        {
            GetWorkSheet(ParseString(sheetName)).Delete();
            return null;
        }

        public CommandData Filter(string tableOrRange, string column, string values)
        {
            try
            {
                ActiveWorksheet.ListObjects[tableOrRange].Range.AutoFilter(int.Parse(column), values.Split(','), XlAutoFilterOperator.xlFilterValues);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Fit(string range)
        {
            try
            {
                ActiveWorksheet.Range[range].Columns.AutoFit();
            }
            catch (Exception) { }
            return null;
        }
        public CommandData FitAll()
        {
            try
            {
                ActiveWorksheet.UsedRange.Columns.AutoFit();
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Formula(string cell, string equation)
        {
            if (TryGetRowCol(cell, out int row, out int col))
                ActiveWorksheet.Cells[row, col].Formula = ParseString(equation); // Remark-cz: Expect starting with '='
            return null;
        }
        public CommandData Formula(string row, string col, string equation)
        {
            try
            {
                ActiveWorksheet.Cells[int.Parse(row), int.Parse(col)].Formula = equation.Trim('"'); // Remark-cz: Expect starting with '='
            }
            catch (Exception) { }
            return null;
        }
        public CommandData InsertRow(string before)
        {
            try
            {
                ActiveWorksheet.Rows[before].Insert();
            }
            catch (Exception){}
            return null;
        }
        public CommandData InsertColumn(string before)
        {
            try
            {
                ActiveWorksheet.Columns[before].Insert();
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Italic(string range)
        {
            try
            {
                ActiveWorksheet.Range[range].Font.Italic = !ActiveWorksheet.Range[range].Font.Italic;
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Italic(string range, string toggle)
        {
            try
            {
                ActiveWorksheet.Range[range].Font.Italic = bool.Parse(toggle);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Merge(string range)
        {
            try
            {
                Application.Range[range].Merge();
            }
            catch (Exception) { }
            return null;
        }
        public CommandData MoveSheetBefore(string sheetName, string otherSheetName)
        {
            try
            {
                GetWorkSheet(ParseString(sheetName)).Move(GetWorkSheet(ParseString(otherSheetName)));
            }
            catch (Exception){}
            return null;
        }
        public CommandData NameRange(string range, string rangeName)
        {
            try
            {
                Globals.ThisAddIn.Application.get_Range(range).Name = ParseString(rangeName);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData NumberFormat(string range, string nameOrFormat)
        {
            try
            {
                Application.Range[range].NumberFormat = nameOrFormat; // Remark-cz: Text or NumberFormat?
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Outline(string range)
        {
            try
            {
                ActiveWorksheet.Range[range].AutoOutline();
            }
            catch (Exception) { }
            return null;
        }
        public CommandData RenameSheet(string newName)
        {
            try
            {
                ActiveWorksheet.Name = newName;
            }
            catch (Exception){}
            return null;
        }
        public CommandData RenameSheet(string originalName, string newName)
        {
            try
            {
                GetWorkSheet(originalName).Name = newName;
            }
            catch (Exception){ throw; }
            return null;
        }
        public CommandData SetCell(string cell, string value)
        {
            if (TryGetRowCol(cell, out int row, out int col))
                ActiveWorksheet.Cells[row, col].Value = ParseValue(value);
            return null;
        }
        public CommandData SetCellName(string cell, string name)
        {
            if (TryGetRowCol(cell, out int row, out int col))
            {
                ActiveWorksheet.Cells[row, col].Name = ParseString(name);
            }
            return null;
        }
        public CommandData SetCellName(string row, string col, string name)
        {
            try
            {
                ActiveWorksheet.Cells[int.Parse(row), int.Parse(col)].Name = ParseString(name);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData SetCell(string row, string col, string value)
        {
            try
            {
                ActiveWorksheet.Cells[int.Parse(row), int.Parse(col)].Value = ParseValue(value);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData SetCellValues(string start, string csv)
        {
            if (TryGetRowCol(start, out int row, out int col))
            {
                var sheet = ActiveWorksheet;
                string[] lines = csv.Split('\n');
                for (int dRow = 0; dRow < lines.Length; dRow++)
                {
                    string[] values = lines[dRow].TrimEnd().Split(',');
                    for (int dCol = 0; dCol < values.Length; dCol++)
                        sheet.Cells[row + dRow, col + dCol].Value = ParseValue(values[dCol]);
                }
            }
            return null;
        }
        public CommandData SetFontColor(string range, string color)
        {
            try
            {
                Application.Range[range].Font.Color = ParseColor(color);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData SetFontSize(string range, string size)
        {
            try
            {
                Application.Range[range].Font.Size = int.Parse(size);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Style(string range, string name)
        {
            try
            {
                Application.Range[range].Style = Application.ActiveWorkbook.Styles[name];
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Width(string range, string width)
        {
            try
            {
                Application.Range[range].ColumnWidth = double.Parse(width);
            }
            catch (Exception) { }
            return null;
        }

        public CommandData Wrap(string range)
        {
            try
            {
                Application.Range[range].WrapText = !Application.Range[range].WrapText;
            }
            catch (Exception) { }
            return null;
        }

        public CommandData Wrap(string range, string toggle)
        {
            try
            {
                Application.Range[range].WrapText = bool.Parse(toggle);
            }
            catch (Exception) { }
            return null;
        }
        #endregion

        #region State Management Routines
        public CommandData Select(string range)
        {
            try
            {
                ActiveWorksheet.Range[range].Select();
            }
            catch (Exception) {}
            return null;
        }
        public CommandData GoToSheet(string sheetName)
        {
            try
            {
                GetWorkSheet(ParseString(sheetName)).Select();
            }
            catch (Exception){}
            return null;
        }
        #endregion

        #region Macros
        public CommandData Apply()
        {
            try
            {
                Application.DoubleClick();
            }
            catch (Exception){}
            return null;
        }
        public CommandData Apply(string range)
        {
            try
            {
                ActiveWorksheet.Range[range].FillDown();    // Remark-cz: Alternatively, we could try AutoFill with the first cell
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Copy()
        {
            try
            {
                Application.Selection.Copy(); // Remark-cz: Pending testing
            }
            catch (Exception){}
            return null;
        }
        public CommandData Duplicate()
        {
            try
            {
                ActiveWorksheet.Copy(ActiveWorksheet);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Fill(string range)
        {
            try
            {
                Application.Selection.AutoFill(ActiveWorksheet.Range[range]);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Fill(string from, string to)
        {
            try
            {
                ActiveWorksheet.Range[from].AutoFill(ActiveWorksheet.Range[to]);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData FillTo(string range, string direction)
        {
            try
            {
                switch (direction.ToLower())
                {
                    case "Up":
                        ActiveWorksheet.Range[range].FillUp();
                        break;
                    case "Left":
                        ActiveWorksheet.Range[range].FillLeft();
                        break;
                    case "Right":
                        ActiveWorksheet.Range[range].FillRight();
                        break;
                    case "Down":
                        ActiveWorksheet.Range[range].FillDown();
                        break;
                    default:
                        break;
                }
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Paste()
        {
            try
            {
                ActiveWorksheet.Paste();
            }
            catch (Exception){}
            return null;
        }
        public CommandData Paste(string range)
        {
            try
            {
                ActiveWorksheet.Paste(ActiveWorksheet.Range[range]);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Save()
        {
            try
            {
                Application.ActiveWorkbook.Save();
            }
            catch (Exception){}
            return null;
        }
        public CommandData Save(string outputFilePath)
        {
            try
            {
                Application.ActiveWorkbook.SaveAs(outputFilePath, XlFileFormat.xlWorkbookNormal);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Sort()
        {
            try
            {
                Application.Selection.Sort(); // Remark-cz: Pending test
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Sort(string range)
        {
            try
            {
                ActiveWorksheet.Range[range].Sort(); // Remark-cz: Pending test
            }
            catch (Exception){}
            return null;
        }
        #endregion

        #region Programming
        public CommandData Evaluate(string scriptPath)
        {
            return null; // Remark-cz: Don't handle this in Excel
        }
        #endregion

        #region Utilities
        public CommandData Random(string range)
        {
            try
            {
                Application.Range[range].Formula = "=RAND()";
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Random(string range, string multiplier)
        {
            try
            {
                Application.Range[range].Formula = $"=RAND()*{multiplier}";
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Random(string range, string from, string to)
        {
            try
            {
                Application.Range[range].Formula = $"=RANDBETWEEN({from}, {to})";
            }
            catch (Exception) { }
            return null;
        }
        #endregion

        #region Finance
        public CommandData ETL(string range, string outputCell)
        {
            return null; // Remark-cz: Not implemented
        }

        public CommandData ETL(string range, string outputCell, string percentage)
        {
            return null; // Remark-cz: Not implemented
        }
        #endregion

        #region Helpers
        private Application Application => Globals.ThisAddIn.Application;
        private Excel.Worksheet ActiveWorksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
        private bool HasWorkSheet(string name)
           => GetWorkSheets().Contains(name);
        private Excel.Worksheet GetWorkSheet(int index)
        {
            try
            {
                return (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[index];
            }
            catch (Exception)
            {
                return null;
            }
        }
        private Excel.Worksheet GetWorkSheet(string name)
        {
            try
            {
                return Globals.ThisAddIn.Application.Worksheets[name];
            }
            catch (Exception)
            {
                return null;
            }
        }
        private string[] GetWorkSheets()
        {
            List<string> sheets = new List<string>();
            foreach (Excel.Worksheet displayWorksheet in Globals.ThisAddIn.Application.Worksheets)
                sheets.Add(displayWorksheet.Name);
            return sheets.ToArray();
        }
        private Excel.Worksheet TryCreateWorksheet(string name)
        {
            if (!HasWorkSheet(name))
            {
                Excel.Worksheet resultSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add();
                resultSheet.Name = name;
                return resultSheet;
            }
            else return GetWorkSheet(name);
        }
        private string GetWorksheetAsCSV(string sheetName)
        {
            return ToCSV(GetWorkSheet(sheetName).UsedRange.Value2);
        }
        #endregion

        #region Value Helpers
        public string ToCSV(object[,] data, string delimiter = ",")
        {
            StringBuilder builder = new StringBuilder();
            for (int row = 0; row < data.GetLength(0); row++)
            {
                for (int col = 0; col < data.GetLength(1); col++)
                {
                    builder.Append(data[row + 1, col + 1]); // For some reason this data from Excel is indexed from 1
                    if (col != data.GetLength(1) - 1)
                        builder.Append(delimiter);
                }
                builder.AppendLine();
            }
            return builder.ToString();
        }
        private bool TryGetRowCol(string cellRef, out int row, out int col)
        {
            var regex = Regex.Match(cellRef, @"([a-zA-Z]+)(\d+)");
            if (regex.Success)
            {
                col = LettersToCol(regex.Groups[1].Value);
                row = int.Parse(regex.Groups[2].Value);

                return true;
            }

            row = -1; col = -1;
            return false;

            int LettersToCol(string letters)
            {
                int value = 0;
                int multiplier = 1;
                foreach (var c in letters.ToLower().Reverse())
                {
                    value += ((int)c - (int)'a' + 1) * multiplier;
                    multiplier *= ((int)'z' - (int)'a' + 1);
                }
                return value;
            }
        }
        private string ParseString(string quotedString)
            => quotedString.Trim('"');
        private object ParseValue(string value)
        {
            return (value.StartsWith("\"") || value.Any(c => !char.IsDigit(c))) ? (object)value.Trim('"') : (object)double.Parse(value);
        }
        private static int ParseColor(string color)
        {
            if (color.StartsWith("#"))
            {
                return System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml(color));
            }
            else if (color.Contains(","))
            {
                string[] parts = color.Split(',');
                return System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(int.Parse(parts[0]), int.Parse(parts[1]), int.Parse(parts[2])));
            }
            else
                return System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromName(color));
        }
        #endregion
    }
}
