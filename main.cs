using OfficeOpenXml;
using System;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        // Example of args: "password", "sheet1", "header1,header2", "row1data1,row1data2", "row2data1,row2data2", "sheet2", ...
        string password = args[0]; // Password for encryption

        using (var package = new ExcelPackage())
        {
            for (int i = 1; i < args.Length;)
            {
                var sheet = package.Workbook.Worksheets.Add(args[i++]); // Sheet name

                var headers = args[i++].Split(','); // Headers for the sheet
                var rowIndex = 1;

                // Add headers
                for (int j = 0; j < headers.Length; j++)
                {
                    sheet.Cells[rowIndex, j + 1].Value = headers[j];
                }

                // Add data rows
                while (i < args.Length && !args[i].StartsWith("sheet"))
                {
                    var rowData = args[i++].Split(',');
                    rowIndex++;
                    for (int j = 0; j < rowData.Length; j++)
                    {
                        sheet.Cells[rowIndex, j + 1].Value = rowData[j];
                    }
                }
            }

            // Apply encryption
            package.Encryption.Password = password;

            // Save the Excel file
            var fileInfo = new FileInfo(@"path\to\your\excel\file.xlsx");
            package.SaveAs(fileInfo);
        }
    }
}
