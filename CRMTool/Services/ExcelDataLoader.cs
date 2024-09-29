using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMTool.Services
{
    internal class ExcelDataLoader
    {
        public static List<Dictionary<string, string>> LoadExcelEntities(string filePath)
        {
            var entities = new List<Dictionary<string, string>>();

            try
            {
                FileInfo fileInfo = new FileInfo(filePath);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming the first sheet

                    int rowCount = worksheet.Dimension.Rows;  // Total number of rows
                    int colCount = worksheet.Dimension.Columns;  // Total number of columns

                    // Get column headers (first row) as keys
                    var headers = new List<string>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        headers.Add(worksheet.Cells[1, col].Text.ToLower());
                    }

                    // Iterate through rows starting from the second row (row 2) since row 1 is the header
                    for (int row = 2; row <= rowCount; row++)
                    {
                        var entityData = new Dictionary<string, string>();
                        bool hasData = false; // Flag to check if there's any data in the row

                        for (int col = 1; col <= colCount; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Text;

                            // Check if the cell is not empty
                            if (!string.IsNullOrWhiteSpace(cellValue))
                            {
                                hasData = true; // Mark that the row has data
                            }

                            // Using header as key, and the cell's value as the value
                            entityData[headers[col - 1]] = cellValue;
                        }

                        // Only add entityData if there's at least one non-empty cell
                        if (hasData)
                        {
                            entities.Add(entityData);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading Excel file: {ex.Message}");
            }

            return entities;
        }

    }
}
