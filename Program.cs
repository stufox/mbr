using System;
using OfficeOpenXml;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;

namespace mbr
{
    class Program
    {
        static int Main(string[] args)
        {
            
            string sheetName = "Raw";
            if (args.Length ==0)
            {
                System.Console.WriteLine("Need to provide an input path on the command line");
                return 1;
            }

            var dir = new DirectoryInfo(args[0]);
            if (dir.Exists)
            { 
                var files = dir.GetFiles("*.csv");
                foreach (var file in files)
                {
                    var csvData = Utils.ReadCSV(file.FullName);   

                    if (csvData.columns > 0)
                    {
                        // create the output array
                        using (var excelPackage = new ExcelPackage())
                        {
                            Utils.TransposeAndClean(excelPackage,sheetName,csvData);
                
                            var worksheet = excelPackage.Workbook.Worksheets[sheetName];
                
                            // Add the graphs/tables

                            ClusterGraph.AddClusterGraph(excelPackage,worksheet);
                            ChangeGraph.AddChangeGraph(excelPackage,worksheet);
                            Linegraph.InsertLineGraph(excelPackage,worksheet);
                            Top20Table.AddTables(excelPackage,worksheet);
                            Top10Table.AddTables(excelPackage,worksheet);
                            Top20Percent.AddTables(excelPackage,worksheet);

                            // write the XLSX file to disk
                            var xlFile = new FileInfo(file.FullName.Replace(".csv",".xlsx"));
                            excelPackage.SaveAs(xlFile);
                        }
                    }
                }
            
            }
            else
            {
                System.Console.WriteLine($"Directory {dir.FullName} does not exist");
            }
            return 0;
        }
    }
}
