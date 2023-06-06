﻿using ExcelCommander.Base.Serialization;
using System.Linq;
using System;
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
                builder.AppendLine($"{method.Name}(${string.Join(",", method.GetParameters().Select(p => $"{p.ParameterType.Name} {p.Name}"))})");
            }
            return builder.ToString();
        }
    }
    public interface ICommander
    {
        #region Reading Routines
        CommandData GetCell(string cell);
        CommandData GetCell(string row, string col);
        CommandData GetCellColor(string cell);
        CommandData GetCellColor(string row, string col);
        CommandData GetCellName(string cell);
        CommandData GetCellName(string row, string col);
        CommandData GetCellFontWeight(string cell);
        CommandData GetCellFontWeight(string row, string col);
        CommandData GetCellValueFormat(string cell);
        CommandData GetCellValueFormat(string row, string col);
        CommandData GetCellValue(string cell);
        CommandData GetCellValue(string row, string col);
        CommandData GetCellFormula(string cell);
        CommandData GetCellFormula(string row, string col);
        CommandData GetCellValues(string cell, string rows, string cols);
        CommandData GetCellValues(string startcell, string endcell);
        CommandData GetCellValues(string range);
        CommandData GetTable(string tableName);
        CommandData GetSheet(string sheetName);
        CommandData GetCurrentSheet();
        CommandData GetSheets();

        CommandData HasSheet(string name);
        CommandData HasTable(string name);
        CommandData HasNamedRange(string name);
        #endregion

        #region Writing Routines
        CommandData CSV(string start, string filename);

        CommandData CreateSheet(string sheetName);
        CommandData MoveSheetBefore(string sheetName, string otherSheetName);

        CommandData CreateTable(string range, string tableName);

        CommandData SetFontWeight(string range, string weight);
        CommandData SetValueFormat(string range, string format);
        CommandData SetColor(string cell, string color);
        CommandData SetColor(string row, string col, string color);
        CommandData SetEquation(string cell, string equation);
        CommandData SetEquation(string row, string col, string equation);
        CommandData SetCell(string cell, string value);
        CommandData SetCell(string row, string col, string value);
        CommandData SetCellName(string cell, string name);
        CommandData SetCellName(string row, string col, string name);
        CommandData SetCellValues(string start, string csv);
        #endregion

        #region State Management Routines
        CommandData GoToSheet(string sheetName);
        #endregion
    }
}