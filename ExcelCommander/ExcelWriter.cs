namespace ExcelCommander
{
    public class ExcelWriter
    {
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
    }
}