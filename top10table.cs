using System;
using OfficeOpenXml;

namespace mbr 
{
    public static class Top10Table
    {

        public static void AddTables(ExcelPackage package, ExcelWorksheet inputsheet)
        {
            string sheetname = "Top10table";
            var sheet = package.Workbook.Worksheets.Copy(inputsheet.Name,sheetname);
            // delete all rows past row 11 (header row plus top 10)
            for (int i=sheet.Dimension.End.Row;i>11;i--)
            {
                sheet.DeleteRow(i);
            }
            


        }

    }

}