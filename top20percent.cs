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
            for (int i=2;i<=sheet.Dimension.End.Row;i++)
            {
                var thisMonthCell = sheet.Cells[i,columns-1].Address;
                var lastMonthCell = sheet.Cells[i,columns-2].Address;
                sheet.Cells[i,columns].Formula = $"{thisMonthCell}-{lastMonthCell}";
                sheet.Cells[i,columns].Style.Numberformat.Format = "\"$\"#,##0.00";
            }
            sheet.Calculate();

            // add a new column for calcuating change in %
            columns = sheet.Dimension.End.Column+1;
            sheet.Cells[1,columns].Value = "Change (%)";
            for (int i=2;i<=sheet.Dimension.End.Row;i++)
            {
                var changeCell = sheet.Cells[i,columns-1].Address;
                //var thisMonthCell = sheet.Cells[i,columns-1].Address;
                var lastMonthCell = sheet.Cells[i,columns-3].Address;
                sheet.Cells[i,columns].Formula = $"{changeCell}/{lastMonthCell}%";
                //sheet.Cells[i,columns].Style.Numberformat.Format = "\"$\"#,##0.00";
            }
            sheet.Calculate();

            // code goes here for sort 
            using (ExcelRange range = sheet.Cells[1,1,sheet.Dimension.End.Row,sheet.Dimension.End.Column])
            {
                //range.Sort(range.Columns-1,Utils.sortDescending);

                System.Console.WriteLine($"would have sorted on row {sheet.Dimension.End.Row} & column {sheet.Dimension.End.Column}");
            }
            // code goes here to recalulate


            // delete all rows past row 21
            /* for (int i=sheet.Dimension.End.Row;i>21;i--)
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
            */


        }

    }

}