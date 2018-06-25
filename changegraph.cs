using System;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;

namespace mbr
{
 
    
    public static class ChangeGraph
    {

        public static void AddChangeGraph(ExcelPackage package, ExcelWorksheet inputsheet, int rows, int columns)
        {
            // Copy the existing worksheet to a new one as we need to make changes
            ExcelWorksheet sheet = package.Workbook.Worksheets.Copy(inputsheet.Name,"Change");
            // Add a new column called Change, and then add a formula into each cell to calculate the change
            // formula is typically M2-L2 or something like that.
            sheet.Cells[1,columns+1].Value = "Change";
            for (int i=2;i<=rows;i++)
            {
                var thisMonthCell = sheet.Cells[i,columns].Address;
                var lastMonthCell = sheet.Cells[i,columns-1].Address;
                sheet.Cells[i,columns+1].Formula = $"{thisMonthCell}-{lastMonthCell}";
                sheet.Cells[i,columns+1].Style.Numberformat.Format = "\"$\"#,##0.00";
            }     
            sheet.Calculate();
            // Sort from highest to lowest.
            using (ExcelRange excelRange = sheet.Cells[1,1,rows,columns+1])
            {
                // sort is zero based, the range isn't so be careful
                // also remember that you're sorting the range, not the entire sheet
                excelRange.Sort(excelRange.Columns-1,Utils.sortDescending);
            }

            // Now remove any rows that are <$100 & >-$100
            // Track how many rows we're removing so we can still maintain how many rows there are
            int adjustment =0;
            for (int i=rows;i>=2;i--)
            {
                double cellValue;

                Double.TryParse(sheet.Cells[i,columns+1].Value.ToString(), out cellValue);
                if (cellValue < Utils.upperSpendLimit && cellValue > Utils.lowerSpendLimit )
                {
                    adjustment++;
                    sheet.DeleteRow(i);
                }
            }
            rows = rows - adjustment;
            // This might be weird, but when you have done a sort on a column with formulas, the sort is "correct"
            // but the formulas in the cells still refer to their source cells. So if you do anything like a recalculate or a sort in Excel
            // things go sideways. So re-enter the formulas to get it to be sorted and correct.
            for (int i=2;i<=rows;i++)
            {
                var thisMonthCell = sheet.Cells[i,columns].Address;
                var lastMonthCell = sheet.Cells[i,columns-1].Address;
                sheet.Cells[i,columns+1].Formula = $"{thisMonthCell}-{lastMonthCell}";

            }  
            
            // Insert our clustered column chart
            var chart = sheet.Drawings.AddChart("Chart",eChartType.ColumnClustered);

            chart.SetSize(1000,800);
            
            // Add the values from each row as a separate series - this is the same as graphing the data in one lump and clicking "Switch row/column" in Excel.
            // because of how we want this to display our label range stays constant as the header cell for the change column (should be "Change")
            for (int i=2;i<=rows;i++)
            {
                var valueRange = ExcelRange.GetAddress(i,columns+1);
                var labelRange = ExcelRange.GetAddress(1,columns+1);
                var series = chart.Series.Add(valueRange,labelRange);
                series.HeaderAddress= new ExcelAddress($"'Change'!A{i}");
            }

            //Formatting
            chart.XAxis.Title.Text ="";
            chart.Title.Text="";
            chart.YAxis.Font.Size = 9;
            chart.XAxis.Font.Size = 9;
            chart.RoundedCorners = false;
            chart.Style = Utils.chartStyle;


        }
    }    
}