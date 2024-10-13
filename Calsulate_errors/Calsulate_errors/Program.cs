using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calsulate_errors
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Define the folder path containing Excel files
            string folderPath = @"C:\Users\katri\Documents\Maija_un_Linda_ZPD\ZPD_Riga_Excel"; // Update this with your folder path

            // Get all Excel files in the folder
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx");

            // Loop through each Excel file
            foreach (var filePath in excelFiles)
            {
                Console.WriteLine($"Processing file: {filePath}");

                FileInfo excelFile = new FileInfo(filePath);

                using (ExcelPackage package = new ExcelPackage(excelFile))
                {
                    // Access the first worksheet
                    var worksheet = package.Workbook.Worksheets[0];

                    // Write headers
                    worksheet.Cells["Z1"].Value = "B (total) STD s";
                    worksheet.Cells["AA1"].Value = "Stjūdena koeficients t (99.9%)";
                    worksheet.Cells["AB1"].Value = "Gadījuma kļūda  δ";
                    worksheet.Cells["AC1"].Value = "Sistemātiskā kļūda θ";
                    worksheet.Cells["AD1"].Value = "Absolūtā kļūda ΔB";
                    worksheet.Cells["AE1"].Value = "Relatīvā kļūda r";

                    // Determine the last row with data in column W
                    int lastRow = worksheet.Dimension.End.Row;

                    // Insert the formulas
                    worksheet.Cells["Z2"].Formula = $"STDEV.S(W2:W{lastRow})"; // B (total) STD s
                    worksheet.Cells["AA2"].Value = 3.291; // Stjūdena koeficients t (99.9%)
                    worksheet.Cells["AB2"].Formula = $"AA2*Z2/SQRT(Q2)"; // Gadījuma kļūda δ
                    worksheet.Cells["AC2"].Value = 0; // Sistemātiskā kļūda θ
                    worksheet.Cells["AD2"].Formula = $"SQRT(AB2^2+AC2^2)"; // Absolūtā kļūda ΔB
                    worksheet.Cells["AE2"].Formula = $"AD2/X2"; // Relatīvā kļūda r

                    // Save the changes to the current file
                    package.Save();
                }

                Console.WriteLine($"File processed: {filePath}");
            }

            Console.WriteLine("All files processed successfully!");
        }

    }
}

