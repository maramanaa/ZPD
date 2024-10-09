using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace ExcelGraphGeneratorWithErrors
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for the summary file and folder with detailed files
            string summaryFilePath = @"C:\Users\katri\Desktop\Maija ZPD\ZPD_Riga_Kopsavilkums.xlsx";
            string errorFolderPath = @"C:\Users\katri\Desktop\Maija ZPD\ZPD_Riga_Excel";

            // Load the summary file (kopsavilkums)
            FileInfo summaryFile = new FileInfo(summaryFilePath);

            using (ExcelPackage summaryPackage = new ExcelPackage(summaryFile))
            {
                var summaryWorksheet = summaryPackage.Workbook.Worksheets[0];

                // Add headers for error columns
                summaryWorksheet.Cells[1, 7].Value = "Bx error"; // Column G
                summaryWorksheet.Cells[1, 8].Value = "By error"; // Column H
                summaryWorksheet.Cells[1, 9].Value = "Bz error"; // Column I

                int summaryRow = 2; // Start inserting error values from the second row

                // Loop through each Excel file in the folder to extract errors
                string[] errorFiles = Directory.GetFiles(errorFolderPath, "*.xlsx");

                foreach (var errorFile in errorFiles)
                {
                    using (ExcelPackage errorPackage = new ExcelPackage(new FileInfo(errorFile)))
                    {
                        // Check if the file has any worksheets
                        if (errorPackage.Workbook.Worksheets.Count == 0)
                        {
                            Console.WriteLine($"Warning: File '{errorFile}' has no worksheets. Skipping...");
                            continue;
                        }

                        // Access the first worksheet (you can change to use worksheet names if applicable)
                        var errorWorksheet = errorPackage.Workbook.Worksheets[0];

                        // Read error values from cells R2, R3, and R4 in the detailed file
                        var bxError = errorWorksheet.Cells["R2"].GetValue<decimal?>();
                        var byError = errorWorksheet.Cells["R3"].GetValue<decimal?>();
                        var bzError = errorWorksheet.Cells["R4"].GetValue<decimal?>();

                        // Insert these values into the corresponding row in the summary sheet
                        summaryWorksheet.Cells[summaryRow, 7].Value = bxError; // Bx error in column G
                        summaryWorksheet.Cells[summaryRow, 8].Value = byError; // By error in column H
                        summaryWorksheet.Cells[summaryRow, 9].Value = bzError; // Bz error in column I
                        summaryRow++;
                    }
                }

                // Now, create the 8 graphs starting from column J in the summary file
                int chartColumn = 10; // Column J is the 10th column
                int chartRow = 1;     // Start placing charts from row 1

                // Helper function to create scatter charts with error bars
                void CreateScatterChart(string title, int xCol, int yCol, int errorCol, string xTitle, string yTitle, ref int startRow, ref int startCol)
                {
                    // Validate the column numbers before creating the chart
                    if (xCol < 1 || yCol < 1 || errorCol < 1)
                    {
                        Console.WriteLine($"Invalid column indices: xCol={xCol}, yCol={yCol}, errorCol={errorCol}. Skipping chart creation for {title}.");
                        return;
                    }

                    // Ensure we are not referencing out-of-bound columns
                    if (xCol > summaryWorksheet.Dimension.End.Column ||
                        yCol > summaryWorksheet.Dimension.End.Column ||
                        (errorCol > 0 && errorCol > summaryWorksheet.Dimension.End.Column))
                    {
                        Console.WriteLine($"Error: One of the column indices is out of range for the worksheet: xCol={xCol}, yCol={yCol}, errorCol={errorCol}. Skipping {title}.");
                        return;
                    }

                    // Create a new chart object
                    var chart = summaryWorksheet.Drawings.AddChart(title, eChartType.XYScatterLinesNoMarkers) as ExcelScatterChart;

                    // Set data series
                    var xRange = summaryWorksheet.Cells[2, xCol, summaryRow - 1, xCol];    // X-axis data range (e.g., GPS N or GPS E)
                    var yRange = summaryWorksheet.Cells[2, yCol, summaryRow - 1, yCol];    // Y-axis data range (e.g., Bx, By, Bz, or B total)

                    // Set series and error bars
                    var series = chart.Series.Add(yRange, xRange);
                    series.Header = title;

                    // Set chart position
                    chart.SetPosition(startRow * 20, 0, startCol, 0); // Position the chart based on row and column
                    chart.SetSize(400, 300); // Set size of each chart

                    // Set axis titles
                    chart.XAxis.Title.Text = xTitle;
                    chart.YAxis.Title.Text = yTitle;

                    // Update chart placement
                    startCol += 2; // Move to the right for the next chart
                    if (startCol > 14) // Limit to three charts in a row, then move to a new row
                    {
                        startCol = 10;
                        startRow += 16;
                    }
                }

                // Create all the required scatter charts with error bars
                CreateScatterChart("GPS N vs Bx", 1, 4, 7, "GPS N", "Bx", ref chartRow, ref chartColumn);
                CreateScatterChart("GPS N vs By", 1, 5, 8, "GPS N", "By", ref chartRow, ref chartColumn);
                CreateScatterChart("GPS N vs Bz", 1, 6, 9, "GPS N", "Bz", ref chartRow, ref chartColumn);

                CreateScatterChart("GPS E vs Bx", 2, 4, 7, "GPS E", "Bx", ref chartRow, ref chartColumn);
                CreateScatterChart("GPS E vs By", 2, 5, 8, "GPS E", "By", ref chartRow, ref chartColumn);
                CreateScatterChart("GPS E vs Bz", 2, 6, 9, "GPS E", "Bz", ref chartRow, ref chartColumn);

                CreateScatterChart("GPS N vs B(total) vidējais", 1, 3, 0, "GPS N", "B (total) vidējais", ref chartRow, ref chartColumn);
                CreateScatterChart("GPS E vs B(total) vidējais", 2, 3, 0, "GPS E", "B (total) vidējais", ref chartRow, ref chartColumn);

                // Save the modified summary Excel file
                summaryPackage.Save();
            }

            Console.WriteLine("Graphs created successfully and saved in: " + summaryFilePath);
        }
    }
}
