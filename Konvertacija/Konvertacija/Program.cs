using OfficeOpenXml;
using System;
using System.IO;
using System.Globalization;


class Program
{
    static void Main(string[] args)
    {
        string folderPath = @"C:\Users\katri\Desktop\Maija ZPD\ZPD Latvija";

        // Check if directory exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine("Folder does not exist.");
            return;
        }

        // Get all TSV files in the folder
        string[] tsvFiles = Directory.GetFiles(folderPath, "*.tsv");

        // Process each TSV file
        foreach (string tsvFile in tsvFiles)
        {
            ConvertTsvToExcel(tsvFile, folderPath);
        }

        Console.WriteLine("Conversion completed!");
    }

    static void ConvertTsvToExcel(string tsvFilePath, string folderPath)
    {
        // Read all lines from the TSV file
        string[] tsvLines = File.ReadAllLines(tsvFilePath);

        // Create a new Excel package
        using (ExcelPackage package = new ExcelPackage())
        {
            // Add a new worksheet
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

            // Process each line (row) of the TSV
            for (int rowIndex = 0; rowIndex < tsvLines.Length; rowIndex++)
            {
                string[] columns = tsvLines[rowIndex].Split('\t'); // Split by tab

                // Process each column (cell) in the row
                for (int colIndex = 0; colIndex < columns.Length; colIndex++)
                {
                    string cellValue = columns[colIndex];

                    // Try to parse the value as a number
                    if (double.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double numericValue))
                    {
                        // If it's a number, store it as a number in the Excel cell
                        worksheet.Cells[rowIndex + 1, colIndex + 1].Value = numericValue;
                    }
                    else
                    {
                        // Otherwise, store it as a string (text)
                        worksheet.Cells[rowIndex + 1, colIndex + 1].Value = cellValue;
                    }
                }
            }

            // Save the Excel file with the same name as the TSV file
            string excelFileName = Path.GetFileNameWithoutExtension(tsvFilePath) + ".xlsx";
            string excelFilePath = Path.Combine(folderPath, excelFileName);

            // Save the Excel file
            FileInfo excelFile = new FileInfo(excelFilePath);
            package.SaveAs(excelFile);
        }

        Console.WriteLine($"Converted: {Path.GetFileName(tsvFilePath)} to {Path.GetFileNameWithoutExtension(tsvFilePath)}.xlsx");
    }
}