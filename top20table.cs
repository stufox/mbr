using System;
using OfficeOpenXml;

namespace mbr 
{
    public static class Top20Table
    {

        public static void AddTables(ExcelPackage package, ExcelWorksheet inputsheet)
        {
            string sheetname = "Top20table";
            var sheet = package.Workbook.Worksheets.Copy(inputsheet.Name,sheetname);
            // delete all rows past row 21
            for (int i=sheet.Dimension.End.Row;i>21;i--)
            {
                sheet.DeleteRow(i);
            }
            // delete all the columns except for the first one and the last two
            // need to delete this many columns starting at column 2 - we should be able to delete column 2 this many times.
            int columnsToDelete = sheet.Dimension.End.Column - 3;

            for (int i=0;i<columnsToDelete;i++)
            {
                sheet.DeleteColumn(2);
            }
            // Add a new column for calculating change in $
            int columns = sheet.Dimension.End.Column+1;
            sheet.Cells[1,columns].Value = "Change ($)";
            sheet.Cells[1,columns].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            sheet.Cells[1,columns].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            for (int i=2;i<=sheet.Dimension.End.Row;i++)
            {
                var thisMonthCell = sheet.Cells[i,columns-1].Address;
                var lastMonthCell = sheet.Cells[i,columns-2].Address;
                sheet.Cells[i,columns].Formula = $"{thisMonthCell}-{lastMonthCell}";
                sheet.Cells[i,columns].Style.Numberformat.Format = "\"$\"#,##0.00";
                sheet.Cells[i,columns].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            }
            sheet.Calculate();

            // add a new column for calcuating change in %
            columns = sheet.Dimension.End.Column+1;
            sheet.Cells[1,columns].Value = "Change (%)";
            sheet.Cells[1,columns].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            sheet.Cells[1,columns].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            for (int i=2;i<=sheet.Dimension.End.Row;i++)
            {
                var changeCell = sheet.Cells[i,columns-1].Address;
                //var thisMonthCell = sheet.Cells[i,columns-1].Address;
                var lastMonthCell = sheet.Cells[i,columns-3].Address;
                sheet.Cells[i,columns].Formula = $"{changeCell}/{lastMonthCell}%";
                //sheet.Cells[i,columns].Style.Numberformat.Format = "\"$\"#,##0.00";
                sheet.Cells[i,columns].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            }
            sheet.Calculate();

        }

    }

}