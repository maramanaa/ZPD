using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart; // to work with charts
using System;
using System.ComponentModel;
using System.IO;

namespace ExcelAutomationWithEPPlus
{
    class Program
    {
        static void Main(string[] args)
        {
            // Folder path 
            string folderPath = @"C:\Users\katri\Desktop\Maija ZPD\ZPD_Riga_Excel";

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;


            try
            {
                // Get all Excel files from the folder
                string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx");

                foreach (string file in excelFiles)
                {
                    Console.WriteLine($"Processing file: {Path.GetFileName(file)}...");

                    // Load the Excel file
                    using (var package = new ExcelPackage(new FileInfo(file)))
                    {
                        // Get the first worksheet in the workbook
                        var worksheet = package.Workbook.Worksheets[0];

                        // Find the last row in column B that contains data
                        int lastRow = worksheet.Dimension.End.Row;

                        // farmulas
                        worksheet.Cells["N2"].Value = "bx";
                        worksheet.Cells["N3"].Value = "by";
                        worksheet.Cells["N4"].Value = "bz";
                        worksheet.Cells["O1"].Value = "AVG";
                        worksheet.Cells["P1"].Value = "STD";
                        worksheet.Cells["Q1"].Value = "N";
                        worksheet.Cells["R1"].Value = "Error";
                        worksheet.Cells["W1"].Value = "B(total)";
                        worksheet.Cells["X1"].Value = "B(total) vidējais";
                        worksheet.Cells["R1"].Value = "Error";
                        worksheet.Cells["S1"].Value = "B(total komp. AVG)";
                        worksheet.Cells["O2"].Formula = $"AVERAGE(B2:B{lastRow})";
                        worksheet.Cells["O3"].Formula = $"AVERAGE(C2:C{lastRow})";
                        worksheet.Cells["O4"].Formula = $"AVERAGE(D2:D{lastRow})";
                        worksheet.Cells["P2"].Formula = $"STDEV.S(B2:B{lastRow})";
                        worksheet.Cells["P3"].Formula = $"STDEV.S(C2:C{lastRow})";
                        worksheet.Cells["P4"].Formula = $"STDEV.S(D2:D{lastRow})";
                        worksheet.Cells["Q2"].Value = lastRow - 1;
                        worksheet.Cells["Q3"].Value = lastRow - 1;
                        worksheet.Cells["Q4"].Value = lastRow - 1;
                        worksheet.Cells["R2"].Formula = $"P2/SQRT(Q2)";
                        worksheet.Cells["R3"].Formula = $"P3/SQRT(Q3)";
                        worksheet.Cells["R4"].Formula = $"P4/SQRT(Q4)";
                        worksheet.Cells["X2"].Formula = $"AVERAGE(W2:W{lastRow})";
                        worksheet.Cells["S2"].Formula = "SQRT((O2)^2+(O3)^2+(O4)^2)";

                        for (int row = 2; row <= lastRow; row++)
                        {
                            // Build the formula for each row
                            string formula = $"=SQRT((B{row})^2+(C{row})^2+(D{row})^2)";

                            // Assign the formula to the W column
                            worksheet.Cells[$"W{row}"].Formula = formula;
                        }



                        // Create the scatter plot (XY chart)
                        var chart = worksheet.Drawings.AddChart("scatterPlot", eChartType.XYScatter);
                        chart.Title.Text = "Laiks pret X komponenti";
                        chart.SetPosition(4, 0, 7, 0);  // Position on the worksheet (Row 10, Column 10)
                        chart.SetSize(600, 400);  // Size of the chart (600x400 pixels)

                        var series = chart.Series.Add(worksheet.Cells[$"B2:B{lastRow}"], worksheet.Cells[$"A2:A{lastRow}"]);
                        series.Header = "Time vs Data";

                        // 1. Create Statistical Histogram for Column B
                        var histogramB = worksheet.Drawings.AddChart("HistogramB", eChartType.Histogram);
                        histogramB.Title.Text = "X komponentes histogramma";
                        histogramB.SetPosition(24, 0, 7, 0); // Adjust position as needed
                        histogramB.SetSize(600, 400); // Adjust size as needed
                        histogramB.Series.Add($"B2:B{lastRow}", $"B2:B{lastRow}"); // Data Range for Column B

                        // 2. Create Statistical Histogram for Column C
                        var histogramC = worksheet.Drawings.AddChart("HistogramC", eChartType.Histogram);
                        histogramC.Title.Text = "Y komponentes histogramma";
                        histogramC.SetPosition(44, 0, 7, 0); // Adjust position as needed
                        histogramC.SetSize(600, 400);
                        histogramC.Series.Add($"C2:C{lastRow}", $"C2:C{lastRow}"); // Data Range for Column C

                        // 3. Create Statistical Histogram for Column D
                        var histogramD = worksheet.Drawings.AddChart("HistogramD", eChartType.Histogram);
                        histogramD.Title.Text = "Z komponentes histogramma";
                        histogramD.SetPosition(64, 0, 7, 0); // Adjust position as needed
                        histogramD.SetSize(600, 400);
                        histogramD.Series.Add($"D2:D{lastRow}", $"D2:D{lastRow}"); // Data Range for Column D

                        // Save the changes
                        package.Save();
                    }

                    Console.WriteLine($"File processed successfully: {Path.GetFileName(file)}");
                }

                Console.WriteLine("All files have been processed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}