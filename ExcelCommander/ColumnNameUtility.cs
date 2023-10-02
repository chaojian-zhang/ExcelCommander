namespace ExcelCommander
{
    public static class ColumnNameUtility
    {
        /// <summary>
        /// Index from 1
        /// </summary>
        public static string GetColumnName(this int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }
    }
}
