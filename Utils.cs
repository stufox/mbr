using System;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml;

namespace mbr
{
    public class Utils
    {
        public const eChartStyle chartStyle = eChartStyle.Style18;
        public const bool sortDescending = true;

        public const double upperSpendLimit = 100.0;
        public const double lowerSpendLimit = -100.0;
        public const int chartHeight = 800;
        public const int chartWidth = 1000;
        public const int defaultFontSize =9;
        public static CSVList ReadCSV(string fileName)
        {
            
            var output = new CSVList();
            Console.WriteLine($"Processing CSV {fileName}");
            using (var streamReader = new StreamReader(fileName))
                {
                    while (!streamReader.EndOfStream)
                    {
                        var line = streamReader.ReadLine();
                        if ((!line.StartsWith("Service Total",StringComparison.CurrentCultureIgnoreCase)) && (!line.StartsWith("LinkedAccount Total",StringComparison.CurrentCultureIgnoreCase)))
                        {
                            output.Content.Add(line.Replace("($)",""));
                            output.rows = (line.Split(",")).Length;
                        }
                    }   
                }
            output.columns = output.Content.Count;
            return output;
        }
        public static void TransposeAndClean(ExcelPackage excelPackage, string sheetName, CSVList csvData)
        {
            var list = csvData.Content;
            var columns = csvData.columns;
            var rows = csvData.rows;
            var worksheet = excelPackage.Workbook.Worksheets.Add(sheetName);
            for (int i=0;i<columns;i++)
            {
                    var splitLine = list[i].Split(",");
                    for (int j=0;j<rows;j++)
                    {
                        // note that excel cell references are one based, not zero based so add one to every reference
                        if (i==0)
                        {
                            // Header column
                            worksheet.Cells[j+1,i+1].Value = splitLine[j];
                        }                       
                        else 
                        { 
                            if (j==0)
                            {
                                // Header row
                                worksheet.Cells[j+1,i+1].Value = DateTime.Parse(splitLine[j]);
                            }
                            else 
                            {     
                                // Data - any empty cells get a zero in it to make the sheet look nicer
                                worksheet.Cells[j+1,i+1].Value = Convert.ToDecimal(String.IsNullOrEmpty(splitLine[j]) ? "0" : splitLine[j]); 
                            }  
                        } 
                    }
                    
            }
            for (int i=worksheet.Dimension.End.Row;i >=1;i--)
                {
                    //System.Console.WriteLine(worksheet.Cells[i,1].Value);
                    string serviceName = worksheet.Cells[i,1].Value.ToString().Trim();
                    if (String.Equals(serviceName, "Total cost") || (String.Equals(serviceName,"Premium Support"))|| (String.Equals(serviceName,"Tax"))||(String.Equals(serviceName,"Refund")))
                    {
                        worksheet.DeleteRow(i);
                    }
                }

                // sort the $ values - the range is from B2 -> the bottom corner of the sheet
                using (ExcelRange excelRange = worksheet.Cells[2,1,worksheet.Dimension.End.Row,worksheet.Dimension.End.Column])
                {
                    // sort is zero based, the range isn't so subtract one to find the last column
                    excelRange.Sort(excelRange.Columns-1,Utils.sortDescending);
                }
            // Align all cells in the middle
            worksheet.Cells[1,1,worksheet.Dimension.Rows,worksheet.Dimension.Columns].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            // Format header dates
            // Header row from B1 onwards gets formatted as date
            worksheet.Cells[1,2,1,worksheet.Dimension.Columns].Style.Numberformat.Format = "mmm-yy";
            worksheet.Cells[1,2,1,worksheet.Dimension.Columns].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            // Format all data cells as currency (B2:End of sheet)
            // set the currency format on the cells - note that $ MUST be enclosed in "" or otherwise it doesn't work properly 
            worksheet.Cells[2,2,worksheet.Dimension.Rows,worksheet.Dimension.Columns].Style.Numberformat.Format = "\"$\"#,##0.00";
            worksheet.Cells[2,2,worksheet.Dimension.Rows,worksheet.Dimension.Columns].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
        }


        public class CSVList
        {
            public CSVList()
            {
                rows =0;
                columns =0;
                Content = new List<string>();
            }
            public int rows {get;set;}
            public int columns {get;set;}

            public List<string>Content {get; set;}

        }
    }
}