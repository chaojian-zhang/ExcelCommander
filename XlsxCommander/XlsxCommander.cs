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

            var methods = CommandMethods;
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
        public CommandData Bold(string range)
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
        public CommandData ClearAll()
        {
            throw new NotImplementedException();
        }
        public CommandData ClearFormat(string range)
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
        public CommandData DeleteColumn(string column)
        {
            throw new NotImplementedException();
        }
        public CommandData DeleteColumns(string columnRange)
        {
            throw new NotImplementedException();
        }
        public CommandData DeleteRow(string row)
        {
            throw new NotImplementedException();
        }
        public CommandData DeleteRows(string rowRange)
        {
            throw new NotImplementedException();
        }
        public CommandData Fit(string range)
        {
            throw new NotFiniteNumberException();
        }
        public CommandData FitAll()
        {
            throw new NotImplementedException();
        }
        public CommandData Formula(string cell, string equation)
        {
            throw new NotImplementedException();
        }
        public CommandData Formula(string row, string col, string equation)
        {
            throw new NotImplementedException();
        }
        public CommandData Italic(string range)
        {
            throw new NotImplementedException();
        }
        public CommandData Italic(string range, string toggle)
        {
            throw new NotImplementedException();
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
        public CommandData Outline(string range)
        {
            throw new NotImplementedException();
        }
        public CommandData RenameSheet(string newName)
        {
            throw new NotImplementedException();
        }
        public CommandData RenameSheet(string originalName, string newName)
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
        public CommandData SetFontColor(string range, string color)
        {
            throw new NotImplementedException();
        }
        public CommandData Width(string range, string width)
        {
            throw new NotImplementedException();
        }
        public CommandData Wrap(string range)
        {
            throw new NotImplementedException();
        }
        public CommandData Wrap(string range, string toggle)
        {
            throw new NotImplementedException();
        }
        #endregion

        #region State Management Routines
        public CommandData Select(string range)
        {
            throw new NotImplementedException();
        }
        public CommandData GoToSheet(string sheetName)
        {
            throw new NotImplementedException();
        }
        #endregion

        #region Macros
        public CommandData Apply()
        {
            throw new NotImplementedException();
        }
        public CommandData Apply(string range)
        {
            throw new NotImplementedException();
        }
        public CommandData Copy()
        {
            throw new NotImplementedException();
        }
        public CommandData Duplicate()
        {
            throw new NotImplementedException();
        }
        public CommandData Fill(string range)
        {
            throw new NotImplementedException();
        }
        public CommandData Fill(string from, string to)
        {
            throw new NotImplementedException();
        }
        public CommandData FillTo(string range, string direction)
        {
            throw new NotImplementedException();
        }
        public CommandData Paste()
        {
            throw new NotImplementedException();
        }
        public CommandData Paste(string range)
        {
            throw new NotImplementedException();
        }
        public CommandData Save()
        {
            throw new NotImplementedException();
        }
        public CommandData Save(string outputFilePath)
        {
            throw new NotImplementedException();
        }
        public CommandData Sort()
        {
            throw new NotImplementedException();
        }
        public CommandData Sort(string range)
        {
            throw new NotImplementedException();
        }
        #endregion

        #region Programming
        public CommandData Evaluate(string scriptPath)
        {
            throw new NotImplementedException();
        }
        #endregion

        #region Utilities
        public CommandData Random(string range)
        {
            throw new NotImplementedException();
        }

        public CommandData Random(string range, string multiplier)
        {
            throw new NotImplementedException();
        }

        public CommandData Random(string range, string from, string to)
        {
            throw new NotImplementedException();
        }
        #endregion

        #region Finance
        public CommandData ETL(string range, string outputCell)
        {
            throw new NotImplementedException();
        }
        public CommandData ETL(string range, string outputCell, string percentage)
        {
            throw new NotImplementedException();
        }
        #endregion
    }
}