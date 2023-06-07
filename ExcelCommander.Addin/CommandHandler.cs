using ExcelCommander.Base.Serialization;
using ExcelCommander.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace ExcelCommander.Addin
{
    internal class CommandHandler: ICommander
    {
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
        #endregion

        #region Entry Point
        internal CommandData Handle(CommandData data)
        {
            return HandleGeneralCommands();

            CommandData HandleGeneralCommands()
            {
                string[] parameters = data.Contents.SplitParameters(true);

                var methods = GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);
                var match = methods.FirstOrDefault(m =>
                    m.Name == parameters[0]
                    && m.GetParameters().Length == parameters.Length - 1
                    && m.ReturnType == typeof(CommandData));
                if (match != null)
                    return (CommandData)match.Invoke(this, parameters.Skip(1).OfType<object>().ToArray());
                else return null;
            }
        }
        #endregion

        #region Reading Routines
        public CommandData GetCell(string cell)
        {
            try
            {
                if (TryGetRowCol(cell, out int row, out int col))
                {
                    return new CommandData()
                    {
                        CommandType = "Value",
                        Contents = ActiveWorksheet.Cells[row, col].Value.ToString()
                    };
                }
            }
            catch (Exception e)
            {
                return Error(e.Message);
            }
            
            return Ok();
        }
        public CommandData GetCell(string row, string col)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCellColor(string cell)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCellColor(string row, string col)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCellName(string cell)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCellName(string row, string col)
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
        public CommandData GetCellValueFormat(string cell)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCellValueFormat(string row, string col)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCellValue(string cell)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCellValue(string row, string col)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCellFormula(string cell)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCellFormula(string row, string col)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCellValues(string cell, string rows, string cols)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCellValues(string startcell, string endcell)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCellValues(string range)
        {
            throw new NotImplementedException();
        }
        public CommandData GetTable(string tableName)
        {
            throw new NotImplementedException();
        }
        public CommandData GetSheet(string sheetName)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCurrentSheet()
        {
            throw new NotImplementedException();
        }
        public CommandData GetSheets()
        {
            throw new NotImplementedException();
        }
        public CommandData HasSheet(string name)
        {
            throw new NotImplementedException();
        }
        public CommandData HasTable(string name)
        {
            return new CommandData()
            {
                CommandType = "Query Result",
                Contents = HasWorkSheet(name).ToString()
            };
        }
        public CommandData HasNamedRange(string name)
        {
            throw new NotImplementedException();
        }
        #endregion

        #region Writing Routines
        public CommandData CSV(string start, string filename)
        {
            try
            {
                return SetCellValues(start, System.IO.File.ReadAllText(filename));
            }
            catch (Exception)
            {
                return null;
            }
        }
        public CommandData CreateSheet(string sheetName)
        {
            TryCreateWorksheet(sheetName);
            return null;
        }
        public CommandData MoveSheetBefore(string sheetName, string otherSheetName)
        {
            try
            {
                GetWorkSheet(sheetName).Move(GetWorkSheet(otherSheetName));
            }
            catch (Exception){}
            return null;
        }
        public CommandData CreateTable(string range, string tableName)
        {
            try
            {
                ActiveWorksheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, Application.get_Range(range), null, XlYesNoGuess.xlYes, null, "TableStyleMedium3").Name = tableName;
            }
            catch (Exception) { }
            return null;
        }
        public CommandData NameRange(string range, string rangeName)
        {
            try
            {
                Globals.ThisAddIn.Application.get_Range(range).Name = rangeName;
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
        public CommandData Bold(string range, string weight)
        {
            try
            {
                Application.Range[range].Style.Font.Bold = bool.Parse(weight);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData SetFontSize(string range, string size)
        {
            try
            {
                Application.Range[range].Style.Font.Size = int.Parse(size);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData SetFontColor(string range, string color)
        {
            try
            {
                Application.Range[range].Style.Font.Color = ParseColor(color);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData Background(string range, string color)
        {
            try
            {
                Application.Range[range].Style.Interior.Color = ParseColor(color);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData SetValueFormat(string range, string format)
        {
            try
            {
                Application.Range[range].NumberFormat = format; // Remark-cz: Text or NumberFormat?
            }
            catch (Exception) { }
            return null;
        }
        public CommandData SetColor(string cell, string color)
        {
            if (TryGetRowCol(cell, out int row, out int col))
                ActiveWorksheet.Cells[row, col].Style.Font.Color = ParseColor(color);
            return null;
        }
        public CommandData SetColor(string row, string col, string color)
        {
            try
            {
                ActiveWorksheet.Cells[int.Parse(row), int.Parse(col)].Style.Font.Color = ParseColor(color);
            }
            catch (Exception) { }
            return null;
        }
        public CommandData SetEquation(string row, string col, string equation)
        {
            try
            {
                ActiveWorksheet.Cells[int.Parse(row), int.Parse(col)].Formula = equation.Trim('"'); // Remark-cz: Expect starting with '='
            }
            catch (Exception) { }
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
                ActiveWorksheet.Cells[row, col].Name = name;
            }
            return null;
        }
        public CommandData SetCellName(string row, string col, string name)
        {
            try
            {
                ActiveWorksheet.Cells[int.Parse(row), int.Parse(col)].Name = name;
            }
            catch (Exception) { }
            return null;
        }
        public CommandData SetEquation(string cell, string equation)
        {
            if (TryGetRowCol(cell, out int row, out int col))
                ActiveWorksheet.Cells[row, col].Formula = equation.Trim('"'); // Remark-cz: Expect starting with '='
            return null;
        }
        public CommandData SetCell(string row, string col, string value)
        {
            try
            {
                ActiveWorksheet.Cells[int.Parse(row), int.Parse(col)].Value = ParseValue(value);
            }
            catch (Exception){}
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
        #endregion

        #region State Management Routines
        public CommandData GoToSheet(string sheetName)
        {
            GetWorkSheet(sheetName).Select();
            return null;
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
        #endregion

        #region Value Helpers
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
        private object ParseValue(string value)
        {
            return (value.StartsWith("\"") || value.Any(c => !char.IsDigit(c))) ? (object)value.Trim('"') : (object)double.Parse(value);
        }
        private static int ParseColor(string color)
        {
            return System.Drawing.ColorTranslator.ToOle((System.Drawing.Color)Enum.Parse(typeof(System.Drawing.Color), color));
        }
        #endregion
    }
}
