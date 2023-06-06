using ExcelCommander.Base.Serialization;
using ExcelCommander.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Reflection;

namespace ExcelCommander.Addin
{
    internal class CommandHandler
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
        public CommandData GoToSheet(string sheetName)
        {
            GetWorkSheet(sheetName).Select();
            return null;
        }
        public CommandData CreateSheet(string sheetName)
        {
            TryCreateWorksheet(sheetName);
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

        #region Helpers
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
        #endregion
    }
}
