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
        const string sheetName = "Spend";
        static int Main(string[] args)
        {
          
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
            /*int columns = csvData.columns;
            int rows = csvData.rows; */
            
            
            // create the output array
            using (var excelPackage = new ExcelPackage())
            {
                int adjustment=0;
                Utils.TransposeAndClean(excelPackage,sheetName,csvData);
                
                var worksheet = excelPackage.Workbook.Worksheets[sheetName];

                int columns = worksheet.Dimension.End.Column;
                //int rows = worksheet.Dimension.End.Row;
                // Post processing - remove some lines.
                for (int i=worksheet.Dimension.End.Row;i >=1;i--)
                {
                    //System.Console.WriteLine(worksheet.Cells[i,1].Value);
                    string serviceName = worksheet.Cells[i,1].Value.ToString().Trim();
                    if (String.Equals(serviceName, "Total cost") || (String.Equals(serviceName,"Premium Support"))|| (String.Equals(serviceName,"Tax"))||(String.Equals(serviceName,"Refund")))
                    {
                        worksheet.DeleteRow(i);
                        adjustment++;
                    }
                }
                
                //rows = rows - adjustment;
                // sort the $ values - the range is from B2 -> the bottom corner of the sheet
                using (ExcelRange excelRange = worksheet.Cells[2,1,worksheet.Dimension.End.Row,columns])
                {
                    // sort is zero based, the range isn't so subtract one to find the last column
                    excelRange.Sort(excelRange.Columns-1,Utils.sortDescending);
                }
                
                // Add the cluster graph
                if (csvData.fileType == "service")
                {
                    ChangeGraph.AddChangeGraph(excelPackage,worksheet);
                    ClusterGraph.AddClusterGraph(excelPackage,worksheet);
                
                }
                if (csvData.fileType == "account")
                {
                    ChangeGraph.AddChangeGraph(excelPackage,worksheet);
                }
                // write the XLSX file to disk
                var xlFile = new FileInfo(file.FullName.Replace(".csv",".xlsx"));
                excelPackage.SaveAs(xlFile);
                
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
