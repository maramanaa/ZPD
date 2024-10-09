using System;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace ExcelReaderWriter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define the folder path containing Excel files
            string folderPath = @"C:\Users\katri\Desktop\Maija ZPD\ZPD_Latvija_Excel";
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx");

            // Define the path for the output Excel file
            string outputFilePath = @"C:\Users\katri\Desktop\Maija ZPD\ZPD_Latvija_kopsavilkums.xlsx";

            using (ExcelPackage outputPackage = new ExcelPackage())
            {
                var outputSheet = outputPackage.Workbook.Worksheets.Add("Results");
                int outputRow = 1;

                // Write headers
                outputSheet.Cells[outputRow, 1].Value = "T2 Values";
                outputSheet.Cells[outputRow, 2].Value = "U2 Values";
                outputRow++;

                // Loop through each Excel file
                foreach (var file in excelFiles)
                {
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];

                        // Read values from T2 and U2
                        var t2Value = worksheet.Cells["T2"].GetValue<decimal?>();
                        var u2Value = worksheet.Cells["U2"].GetValue<decimal?>();

                        // Write the values to the new sheet
                        outputSheet.Cells[outputRow, 1].Value = t2Value;
                        outputSheet.Cells[outputRow, 2].Value = u2Value;
                        outputRow++;
                    }
                }

                // Save the output Excel file
                File.WriteAllBytes(outputFilePath, outputPackage.GetAsByteArray());
            }

            Console.WriteLine("Data extraction complete! Output saved at: " + outputFilePath);

        }
    }
}
