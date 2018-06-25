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
        static void Main(string[] args)
        {
          
            var csvData = Utils.ReadCSV(@"/Users/stufox/Documents/test/costsbyservice.csv");            
            List<string> list = csvData.Content;
            int columns = csvData.columns;
            int rows = csvData.rows;

            
            // create the output array
            using (var excelPackage = new ExcelPackage())
            {
                int adjustment=0;
                var worksheet = excelPackage.Workbook.Worksheets.Add("Spend");
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
                            // set the currency format on the cells - note that $ MUST be enclosed in " " or otherwise it doesn't work properly          
                            worksheet.Cells[j+1,i+1].Style.Numberformat.Format = "\"$\"#,##0.00";
                            worksheet.Cells[j+1,i+1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }                      
                    }
                    

                }
                // Post processing - remove some lines.
                for (int i=rows;i >=1;i--)
                {
                    //System.Console.WriteLine(worksheet.Cells[i,1].Value);
                    string serviceName = worksheet.Cells[i,1].Value.ToString().Trim();
                    if (String.Equals(serviceName, "Total cost") || (String.Equals(serviceName,"Premium Support"))|| (String.Equals(serviceName,"Tax"))||(String.Equals(serviceName,"Refund")))
                    {
                        worksheet.DeleteRow(i);
                        adjustment++;
                    }
                }
                rows = rows - adjustment;
                // sort the $ values - the range is from B2 -> the bottom corner of the sheet
                using (ExcelRange excelRange = worksheet.Cells[2,1,rows,columns])
                {
                    // sort is zero based, the range isn't so subtract one to find the last column
                    excelRange.Sort(excelRange.Columns-1,Utils.sortDescending);
                }
                
                // Add the cluster graph
                if (csvData.fileType == "service")
                {
                    ChangeGraph.AddChangeGraph(excelPackage,worksheet,rows,columns);
                    ClusterGraph.AddClusterGraph(excelPackage,worksheet,rows,columns);
                
                }
                // write the XLSX file to disk
                var xlFile = new FileInfo(@"/users/stufox/Documents/test/testxl.xlsx");
                excelPackage.SaveAs(xlFile);
            }


        }
    }
}
