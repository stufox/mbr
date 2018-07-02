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
        public const int topN = 5;
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
                            var newLine = line.Replace("($)","");

                            output.Content.Add(newLine);
                            var strings = line.Split(",");
                            output.rows = strings.Length;
                        }
                        if (line.StartsWith("Service Total"))
                        {
                            output.fileType = "service";
                        }
                        if (line.StartsWith("LinkedAccount Total"))
                        {
                            output.fileType = "account";
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
                            worksheet.Cells[j+1,i+1].Value = splitLine[j];
                        }                       
                        if ((j==0) && (i>0))
                        {
                            // Header row from B1 onwards gets formatted as date
                            worksheet.Cells[j+1,i+1].Value = DateTime.Parse(splitLine[j]);
                            worksheet.Cells[j+1,i+1].Style.Numberformat.Format = "mmm-yy";                           
                        }
                        if ((j>0)&&(i>0))
                        {     
                             // any empty cells get a zero in it to make the sheet look nicer
                            worksheet.Cells[j+1,i+1].Value = Convert.ToDecimal(String.IsNullOrEmpty(splitLine[j]) ? "0" : splitLine[j]); 
                            // set the currency format on the cells - note that $ MUST be enclosed in "" or otherwise it doesn't work properly          
                            worksheet.Cells[j+1,i+1].Style.Numberformat.Format = "\"$\"#,##0.00";
                            worksheet.Cells[j+1,i+1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }                      
                    }
        }
        }


        public class CSVList
        {
           
            public CSVList()
            {
                rows =0;
                columns =0;
                Content = new List<string>();
                fileType = "";
            }
            public int rows {get;set;}
            public int columns {get;set;}

            public List<string>Content {get; set;}
            public string fileType{get;set;}

        }

    }
}