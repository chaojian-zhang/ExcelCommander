using ExcelCommander.Base.Serialization;
using System.Linq;
using System.Text;

namespace ExcelCommander.Base
{
    public static class CommanderHelper
    {
        public static string GetHelpString()
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("Available commands: ");
            foreach (var method in typeof(ICommander).GetMethods())
            {
                builder.AppendLine($"{method.Name}({string.Join(", ", method.GetParameters().Select(p => $"{p.ParameterType.Name} {p.Name}"))})");
            }
            return builder.ToString();
        }
    }
    public interface ICommander
    {
        #region Reading Routines
        CommandData Get(string range);
        CommandData GetCell(string cell);
        CommandData GetCell(string row, string col);
        CommandData GetCellColor(string cell);
        CommandData GetCellColor(string row, string col);
        CommandData GetCellFontWeight(string cell);
        CommandData GetCellFontWeight(string row, string col);
        CommandData GetCellFormula(string cell);
        CommandData GetCellFormula(string row, string col);
        CommandData GetCellName(string cell);
        CommandData GetCellName(string row, string col);
        CommandData GetCellValue(string cell);
        CommandData GetCellValue(string row, string col);
        CommandData GetNumberFormat(string range);
        CommandData GetNumberFormat(string row, string col);
        CommandData GetCellValues(string cell, string rows, string cols);
        CommandData GetCellValues(string range);
        CommandData GetCurrentSheet();
        CommandData GetSheet(string sheetName);
        CommandData GetSheets();
        CommandData GetTable(string tableName);
        CommandData HasNamedRange(string name);
        CommandData HasSheet(string name);
        CommandData HasTable(string name);
        #endregion

        #region Writing Routines
        CommandData Align(string range, string option);
        CommandData Background(string range, string color);
        CommandData Bold(string range);
        CommandData Bold(string range, string toggle);
        CommandData Border(string range, string weight);
        CommandData Cell(string range, string value);
        CommandData Clear(string range);
        CommandData ClearAll();
        CommandData ClearFormat(string range);
        CommandData Color(string range, string color);
        CommandData Color(string row, string col, string color);
        CommandData CreateSheet(string sheetName);
        CommandData CreateTable(string range, string tableName);
        CommandData CSV(string start, string filename);
        CommandData DeleteColumn(string column);
        CommandData DeleteColumns(string columnRange);
        CommandData DeleteRow(string row);
        CommandData DeleteRows(string rowRange);
        CommandData DeleteSheet(string sheetName);
        CommandData Filter(string tableOrRange, string column, string values);
        CommandData Fit(string range);
        CommandData FitAll();
        CommandData Formula(string cell, string equation);
        CommandData Formula(string row, string col, string equation);
        CommandData Italic(string range);
        CommandData Italic(string range, string toggle);
        CommandData Merge(string range);
        CommandData MoveSheetBefore(string sheetName, string otherSheetName);
        CommandData NameRange(string range, string name);
        CommandData NumberFormat(string range, string nameOrFormat);
        CommandData Outline(string range);
        CommandData RenameSheet(string newName);
        CommandData RenameSheet(string originalName, string newName);
        CommandData SetCell(string cell, string value);
        CommandData SetCell(string row, string col, string value);
        CommandData SetCellName(string cell, string name);
        CommandData SetCellName(string row, string col, string name);
        CommandData SetCellValues(string start, string csv);
        CommandData SetFontColor(string range, string color);
        CommandData SetFontSize(string range, string size);
        CommandData Style(string range, string name);
        CommandData Width(string range, string width);
        CommandData Wrap(string range);
        CommandData Wrap(string range, string toggle);
        #endregion

        #region State Management Routines
        CommandData Select(string range);
        CommandData GoToSheet(string sheetName);
        #endregion

        #region Macro
        CommandData Apply();
        CommandData Apply(string range);
        CommandData Copy();
        CommandData Duplicate();
        CommandData Fill(string range); // Remark-cz: Alias to Appy()
        CommandData Fill(string from, string to);
        CommandData FillTo(string range, string direction);
        CommandData InsertRow(string before);
        CommandData InsertColumn(string before);
        CommandData Paste();
        CommandData Paste(string range);
        CommandData Save();
        CommandData Save(string outputFilePath);
        CommandData Sort();
        CommandData Sort(string range);
        #endregion

        #region Programming
        CommandData Evaluate(string scriptPath);
        #endregion

        #region Utilities
        CommandData Random(string range);
        CommandData Random(string range, string multiplier);
        CommandData Random(string range, string from, string to);
        #endregion

        #region Finance
        CommandData ETL(string range, string outputCell);
        CommandData ETL(string range, string outputCell, string percentage);
        #endregion
    }
}
