using System;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml.Drawing.Chart;

namespace mbr
{
    public class Utils
    {
        public const eChartStyle chartStyle = eChartStyle.Style18;
        public const bool sortDescending = true;

        public const double upperSpendLimit = 100.0;
        public const double lowerSpendLimit = -100.0;
        public static CSVList ReadCSV(string fileName)
        {
            
            var output = new CSVList();

            
            Console.WriteLine("Processing CSV");
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

            
            System.Console.WriteLine($"The array is {output.rows} x {output.columns}");

            
            return output;

            
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