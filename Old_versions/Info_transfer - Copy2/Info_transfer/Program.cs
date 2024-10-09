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
            string folderPath = @"C:\Users\katri\Desktop\Maija ZPD\ZPD_Riga_Excel";
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx");

            // Define the path for the output Excel file
            string outputFilePath = @"C:\Users\katri\Desktop\Maija ZPD\ZPD_Riga_Kopsavilkums.xlsx";

            using (ExcelPackage outputPackage = new ExcelPackage())
            {
                var outputSheet = outputPackage.Workbook.Worksheets.Add("Results");
                int outputRow = 1;

                // Write headers
                outputSheet.Cells[outputRow, 1].Value = "T2 Values";
                outputSheet.Cells[outputRow, 2].Value = "U2 Values";
                outputSheet.Cells[outputRow, 3].Value = "S2 Values";  // New column for S2 values
                outputRow++;

                // Loop through each Excel file
                foreach (var file in excelFiles)
                {
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];

                        // Read values from T2, U2, and S2, and handle different data types
                        var t2Value = GetDecimalValue(worksheet.Cells["T2"]);
                        var u2Value = GetDecimalValue(worksheet.Cells["U2"]);
                        var s2Value = GetDecimalValue(worksheet.Cells["S2"]);

                        // Write the values to the new sheet, ensuring valid values are placed
                        outputSheet.Cells[outputRow, 1].Value = t2Value.HasValue ? t2Value.Value : (object)"Invalid Data";
                        outputSheet.Cells[outputRow, 2].Value = u2Value.HasValue ? u2Value.Value : (object)"Invalid Data";
                        outputSheet.Cells[outputRow, 3].Value = s2Value.HasValue ? s2Value.Value : (object)"Invalid Data";
                        outputRow++;
                    }
                }

                // Save the output Excel file
                File.WriteAllBytes(outputFilePath, outputPackage.GetAsByteArray());
            }

            Console.WriteLine("Data extraction complete! Output saved at: " + outputFilePath);
        }

        // Helper function to safely extract decimal values from a cell
        private static decimal? GetDecimalValue(ExcelRange cell)
        {
            try
            {
                if (cell.Value != null && decimal.TryParse(cell.Text, out var result))
                {
                    return result;
                }
                return null;  // Return null if the cell is empty or contains non-numeric text
            }
            catch
            {
                return null;  // Return null if any other exception occurs during conversion
            }

        }
    }
}
