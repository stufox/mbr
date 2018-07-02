using System;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;

namespace mbr
{
    public static class ClusterGraph
    {
        public static void AddClusterGraph(ExcelPackage package, ExcelWorksheet inputsheet) //, int rows, int columns)
        {
            
            string sheetName = "Spend";
            var sheet = package.Workbook.Worksheets.Copy(inputsheet.Name,sheetName);
            // now find the position where $$ is less than the spendLimit value for graphing 
            // (We don't want to graph small numbers)
            Double cellValue = sheet.Cells[sheet.Dimension.End.Row, sheet.Dimension.End.Column].GetValue<double>();
            int rows = sheet.Dimension.End.Row;
            int columns = sheet.Dimension.End.Column;
            int position = sheet.Dimension.End.Row;
            while (cellValue < Utils.upperSpendLimit)
            {
                position--;
                cellValue = sheet.Cells[position,sheet.Dimension.End.Column].GetValue<double>();
            }
            

            // Add the chart
            var barChart = sheet.Drawings.AddChart("Chart",eChartType.ColumnStacked) as ExcelBarChart;
            barChart.SetSize(1000,800);

                
            // Notes:   for Series.Add() first series is values - in this case will be Bx to Mx
            // second series is labels for X Axis - should be just B1 to M1
            for (int i=2;i<=position;i++)
            {
                var valueRange = ExcelRange.GetAddress(i,2,i,sheet.Dimension.End.Column);
                var labelRange = ExcelRange.GetAddress(1,2,1,sheet.Dimension.End.Column);
                var series = barChart.Series.Add(valueRange,labelRange);
                series.HeaderAddress = new ExcelAddress($"'{sheetName}'!A{i}");   
            } 
                
            barChart.Title.Text="";
            barChart.YAxis.Font.Size = Utils.defaultFontSize;
            barChart.RoundedCorners = false;
            barChart.Style = Utils.chartStyle;
            
        }


    }

}