using System;
using OfficeOpenXml;

namespace mbr 
{
    public static class Top20Percent
    {

        public static void AddTables(ExcelPackage package, ExcelWorksheet inputsheet)
        {
            string sheetname = "Top20percent";
            var sheet = package.Workbook.Worksheets.Copy(inputsheet.Name,sheetname);

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
                // bugfix July 18
                // Need to not do a divide by zero error
                // So check for last month cell being zero
                

                var lastMonthCell = sheet.Cells[i,columns-3].Address;
                if (sheet.Cells[lastMonthCell].GetValue<double>() != 0) 
                {
                    sheet.Cells[i,columns].Formula = $"{changeCell}/{lastMonthCell}%";

                }
                else
                {
                    sheet.Cells[i,columns].Value = 0.0;
                    
                }
                sheet.Cells[i,columns].Style.Numberformat.Format = "0.00";
                sheet.Cells[i,columns].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            }
            sheet.Calculate();

            // code goes here for sort 
            using (ExcelRange range = sheet.Cells[1,1,sheet.Dimension.End.Row,sheet.Dimension.End.Column])
            {
                range.Sort(range.Columns-1,Utils.sortDescending);

            }

            // now re-enter all the formulas and recalculate
            columns = sheet.Dimension.End.Column;
            for (int i=2;i<=sheet.Dimension.End.Row;i++)
            {
                var thisMonthCell = sheet.Cells[i,columns-2].Address;
                var lastMonthCell = sheet.Cells[i,columns-3].Address;
                sheet.Cells[i,columns-1].Formula = $"{thisMonthCell}-{lastMonthCell}";
                sheet.Cells[i,columns-1].Style.Numberformat.Format = "\"$\"#,##0.00";
            }
            sheet.Calculate();

            for (int i=2;i<=sheet.Dimension.End.Row;i++)
            {
                var changeCell = sheet.Cells[i,columns-1].Address;
                var lastMonthCell = sheet.Cells[i,columns-3].Address;

                    sheet.Cells[i,columns].Formula = $"{changeCell}/{lastMonthCell}%";
                    sheet.Cells[i,columns].Style.Numberformat.Format = "0.00";

            }
            sheet.Calculate();




            // delete all rows past row 21
            for (int i=sheet.Dimension.End.Row;i>21;i--)
            {
                sheet.DeleteRow(i);
            }
            // delete all the columns except for the first one and the last five (so we see three month spending trend)
            // need to delete this many columns starting at column 2 - we should be able to delete column 2 this many times.
            int columnsToDelete = sheet.Dimension.End.Column - 6;

            for (int i=0;i<columnsToDelete;i++)
            {
                sheet.DeleteColumn(2);
            }
            
            sheet.Cells.AutoFitColumns();

        }

    }

}