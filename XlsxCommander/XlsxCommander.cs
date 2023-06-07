using ExcelCommander.Base;
using ExcelCommander.Base.Serialization;
using System.Reflection;

namespace XlsxCommander
{
    public class XlsxCommander: ICommander
    {
        #region Construction
        public static XlsxCommander Start(string outputFile)
            => new XlsxCommander(outputFile);
        public XlsxCommander(string outputFile)
        {
            OutputFile = outputFile;
        }
        public string OutputFile { get; }
        #endregion

        #region State

        #endregion

        #region Parsing Methods
        public void Execute(string[] commands, bool interpretIfNull = true)
        {
            if (commands == null && interpretIfNull)
            {
                while (true)
                {
                    Console.Write("> ");
                    string input = Console.ReadLine();
                    ExecuteCommand(input);
                }
            }
            else
            {
                foreach (var command in commands)
                    ExecuteCommand(command);
            }
        }
        public void ExecuteCommand(string command)
        {
            if (command == "Help")
                Console.WriteLine(CommanderHelper.GetHelpString());
            else
                EvaluateCommand(command);
        }
        #endregion

        #region Interface
        internal void EvaluateCommand(string command)
        {
            string[] parameters = command.SplitParameters(true);

            var methods = GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);
            var match = methods.FirstOrDefault(m =>
                m.Name == parameters[0]
                && m.GetParameters().Length == parameters.Length - 1
                && m.ReturnType == typeof(CommandData)); // Remark-cz: ExcelWriter implements the same CommandData return type interface but generally do not return anything as messages (but may return as payloads) and preferrably prints out the results
            match?.Invoke(this, parameters.Skip(1).OfType<object>().ToArray());
        }
        #endregion

        #region Speciaty Functions
        public void Spawn()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            Microsoft.Office.Interop.Excel.Workbooks workbooks = null;
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            workbooks = excelApp.Workbooks;
            workbook = workbooks.Add(1);
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
            excelApp.Visible = true;
            worksheet.Cells[1, 1] = "Value1";
            worksheet.Cells[1, 2] = "Value2";
            worksheet.Cells[1, 3] = "Addition";
            worksheet.Cells[2, 1] = 1;
            worksheet.Cells[2, 2] = 2;
            worksheet.Cells[2, 3].Formula = "=SUM(A2,B2)";
        }
        #endregion

        #region Reading Routines
        public CommandData Get(string range)
        {
            throw new NotImplementedException();
        }
        public CommandData GetCell(string cell)
        {
            throw new NotImplementedException();
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
            throw new NotImplementedException();
        }
        public CommandData HasNamedRange(string name)
        {
            throw new NotImplementedException();
        }
        #endregion

        #region Writing Routines
        public CommandData Align(string range, string option)
        {
            throw new NotImplementedException();
        }
        public CommandData Background(string range, string color)
        {
            throw new NotImplementedException();
        }
        public CommandData Bold(string range, string weight)
        {
            throw new NotImplementedException();
        }
        public CommandData Border(string range, string weight)
        {
            throw new NotImplementedException();
        }
        public CommandData Cell(string range, string value)
        {
            throw new NotImplementedException();
        }
        public CommandData Clear(string range)
        {
            throw new NotImplementedException();
        }
        public CommandData ClearFormat(string range)
        {
            throw new NotImplementedException();
        }
        public CommandData CreateSheet(string sheetName)
        {
            throw new NotImplementedException();
        }
        public CommandData CreateTable(string range, string tableName)
        {
            throw new NotImplementedException();
        }
        public CommandData CSV(string start, string filename)
        {
            throw new NotImplementedException();
        }
        public CommandData Fit(string range)
        {
            throw new NotFiniteNumberException();
        }
        public CommandData Merge(string range)
        {
            throw new NotImplementedException();
        }
        public CommandData MoveSheetBefore(string sheetName, string otherSheetName)
        {
            throw new NotImplementedException();
        }
        public CommandData NameRange(string range, string name)
        {
            throw new NotImplementedException();
        }
        public CommandData SetValueFormat(string range, string format)
        {
            throw new NotImplementedException();
        }
        public CommandData SetFontSize(string range, string size)
        {
            throw new NotImplementedException();
        }
        public CommandData SetCell(string cell, string value)
        {
            throw new NotImplementedException();
        }
        public CommandData SetCell(string row, string col, string value)
        {
            throw new NotImplementedException();
        }
        public CommandData SetCellName(string cell, string name)
        {
            throw new NotImplementedException();
        }
        public CommandData SetCellName(string row, string col, string name)
        {
            throw new NotImplementedException();
        }
        public CommandData SetCellValues(string start, string csv)
        {
            throw new NotImplementedException();
        }
        public CommandData Color(string range, string color)
        {
            throw new NotImplementedException();
        }
        public CommandData Color(string row, string col, string color)
        {
            throw new NotImplementedException();
        }
        public CommandData SetEquation(string cell, string equation)
        {
            throw new NotImplementedException();
        }
        public CommandData SetEquation(string row, string col, string equation)
        {
            throw new NotImplementedException();
        }
        public CommandData SetFontColor(string range, string color)
        {
            throw new NotImplementedException();
        }
        #endregion

        #region State Management Routines
        public CommandData GoToSheet(string sheetName)
        {
            throw new NotImplementedException();
        }
        #endregion
    }
}