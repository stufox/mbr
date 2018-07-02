using System;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;

namespace mbr
{
 
    
    public static class ChangeGraph
    {

        public static void AddChangeGraph(ExcelPackage package, ExcelWorksheet inputsheet) 
        {
            string sheetName = "Change";
            // Copy the existing worksheet to a new one as we need to make changes
            ExcelWorksheet sheet = package.Workbook.Worksheets.Copy(inputsheet.Name,sheetName);
            // Add a new column called Change, and then add a formula into each cell to calculate the change
            // formula is typically M2-L2 or something like that.

            int columns = sheet.Dimension.End.Column+1;
            int rows = sheet.Dimension.End.Row;
            sheet.Cells[1,columns].Value = "Change";


             // NOTE TO SELF: You have to use a variable that won't change when you're adding entries to a spreadsheet 
            // If you just refer to the end of the spreadsheet, that tends to move when you add stuff.           
            
            for (int i=2;i<=rows;i++)
            {
                var thisMonthCell = sheet.Cells[i,columns-1].Address;
                var lastMonthCell = sheet.Cells[i,columns-2].Address;
                sheet.Cells[i,columns].Formula = $"{thisMonthCell}-{lastMonthCell}";
                sheet.Cells[i,columns].Style.Numberformat.Format = "\"$\"#,##0.00";
            
            }     
            
            sheet.Calculate();
            
            // Sort from highest to lowest.
            using (ExcelRange excelRange = sheet.Cells[1,1,sheet.Dimension.End.Row,sheet.Dimension.End.Column])
            {
                // sort is zero based, the range isn't so be careful
                // also remember that you're sorting the range, not the entire sheet
                excelRange.Sort(excelRange.Columns-1,Utils.sortDescending);
            }

            // Now remove any rows that are <$100 & >-$100

            for (int i=sheet.Dimension.End.Row;i>=2;i--)
            {
                double cellValue;

                Double.TryParse(sheet.Cells[i,sheet.Dimension.End.Column].Value.ToString(), out cellValue);
                if (cellValue < Utils.upperSpendLimit && cellValue > Utils.lowerSpendLimit )
                {

                    sheet.DeleteRow(i);
                }
            }

            // This might be weird, but when you have done a sort on a column with formulas, the sort is "correct"
            // but the formulas in the cells still refer to their source cells. So if you do anything like a recalculate or a sort in Excel
            // things go sideways. So re-enter the formulas to get it to be sorted and correct.
            for (int i=2;i<=sheet.Dimension.End.Row;i++)
            {
                var thisMonthCell = sheet.Cells[i,sheet.Dimension.End.Column-1].Address;
                var lastMonthCell = sheet.Cells[i,sheet.Dimension.End.Column-2].Address;
                sheet.Cells[i,columns].Formula = $"{thisMonthCell}-{lastMonthCell}";

            }  
            
            // Insert our clustered column chart
            var chart = sheet.Drawings.AddChart("Chart",eChartType.ColumnClustered);

            chart.SetSize(Utils.chartWidth,Utils.chartHeight);
            
            // Add the values from each row as a separate series - this is the same as graphing the data in one lump and clicking "Switch row/column" in Excel.
            // because of how we want this to display our label range stays constant as the header cell for the change column (should be "Change")
            for (int i=2;i<=sheet.Dimension.End.Row;i++)
            {
                var valueRange = ExcelRange.GetAddress(i,sheet.Dimension.End.Column);
                var labelRange = ExcelRange.GetAddress(1,sheet.Dimension.End.Column);
                var series = chart.Series.Add(valueRange,labelRange);
                series.HeaderAddress= new ExcelAddress($"'{sheetName}'!A{i}");
            }

            //Formatting
            chart.XAxis.Title.Text ="";
            chart.Title.Text="";
            chart.YAxis.Font.Size = Utils.defaultFontSize;
            chart.XAxis.Font.Size = Utils.defaultFontSize;
            chart.RoundedCorners = false;
            chart.Style = Utils.chartStyle;


        }
    }    
}