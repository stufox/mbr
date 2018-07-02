using System;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System.Collections.Generic;


namespace mbr
{
    class Linegraph
    {
        public static void InsertLineGraph(ExcelPackage package, ExcelWorksheet inputSheet)
        {
            int seriesStartColumn =2; // i.e. first column of data, column 1 is header

            int seriesHeaderRow = 1;
            string sheetname = "Trend";
            var graphRows = new List<int>(){6,11}; // This list holds the rows for the graphs - 6 is for top 5, 11 is for top 10 (header is one row)
            var sheet = package.Workbook.Worksheets.Copy(inputSheet.Name,sheetname);

            int counter=0;
            // create a top 10 graph & a top 5 graph
            foreach (var graphLimit in graphRows)
            {

                var chart = sheet.Drawings.AddChart($"Top{graphLimit}",eChartType.Line);
                chart.SetSize(Utils.chartWidth,Utils.chartHeight);
                chart.Style = Utils.chartStyle;
                chart.Title.Text="";
                chart.YAxis.Font.Size = Utils.defaultFontSize;
                chart.XAxis.Font.Size = Utils.defaultFontSize;
                chart.RoundedCorners = false;
                
                chart.SetPosition(Utils.chartHeight*counter,0);
                for (int i=2;i<=graphLimit;i++)
                {

                    var valueRange = ExcelRange.GetAddress(i,seriesStartColumn,i,sheet.Dimension.End.Column);
                    var labelRange = ExcelRange.GetAddress(seriesHeaderRow,seriesStartColumn,seriesHeaderRow,sheet.Dimension.End.Column);
                    var series = chart.Series.Add(valueRange,labelRange);
                    series.HeaderAddress = new ExcelAddress($"'{sheetname}'!A{i}");
                }
                counter++;
            }

   
        }
        
    }

}