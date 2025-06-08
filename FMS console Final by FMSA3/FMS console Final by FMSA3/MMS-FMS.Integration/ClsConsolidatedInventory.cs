using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;
using OfficeOpenXml;

class Program
{
    static async Task Main(string[] args)
    {
        try
        {
            string folderPath = @"E:\ConsoleApp\FMS console Final by FMSA3\FMS console Final by FMSA3\excelworkbook\files";
            string outputFile = @"E:\ConsoleApp\FMS console Final by FMSA3\FMS console Final by FMSA3\excelworkbook\consolidated_inventory.xlsx";

            // Create an instance of InventoryConsolidator
            var consolidator = new InventoryConsolidator();

            // Start the consolidation process
            await consolidator.ConsolidateInventoryDataAsync(folderPath, outputFile);

            // Print message when the process is complete
            Console.WriteLine("Consolidation complete!");
        }
        catch (Exception ex)
        {
            // Catch any unhandled exception and display the error message
            Console.WriteLine($"An error occurred: {ex.Message}");
            Console.WriteLine(ex.StackTrace);  // Optionally display the full stack trace for debugging
        }
    }
}

class InventoryConsolidator
{
    public async Task ConsolidateInventoryDataAsync(string folderPath, string outputFile)
    {
        try
        {
            var files = Directory.GetFiles(folderPath, "*.xlsx");
            var tasks = new List<Task<List<InventoryItem>>>();

            // Process each file concurrently
            foreach (var file in files)
            {
                tasks.Add(Task.Run(() => ProcessFile(file)));
            }

            // Await all tasks to complete
            var results = await Task.WhenAll(tasks);

            // Consolidate all results into one list
            var allItems = results.SelectMany(result => result).ToList();

            // Write the consolidated data to a new file
            WriteConsolidatedData(allItems, outputFile);
        }
        catch (Exception ex)
        {
            // Handle any error that occurs in the consolidation process
            Console.WriteLine($"An error occurred while consolidating the data: {ex.Message}");
            Console.WriteLine(ex.StackTrace);  // Optionally display the full stack trace
        }
    }

    private List<InventoryItem> ProcessFile(string filePath)
    {
        var items = new List<InventoryItem>();

        try
        {
            // Read the Excel file using EPPlus
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];  // Assuming the first sheet
                int rowCount = worksheet.Dimension.Rows;
                int columnCount = worksheet.Dimension.Columns;

                // Step 1: Find the header row (you may need to adjust this depending on file structure)
                int headerRow = FindHeaderRow(worksheet, rowCount);

                // Step 2: Identify relevant columns by header names
                var columnIndices = GetColumnIndices(worksheet, headerRow, columnCount);

                // Step 3: Process each data row
                for (int row = 2; row <= rowCount; row++)
                {
                    var item = new InventoryItem
                    {
                        ItemId = worksheet.Cells[row, 1].Text,  // Assuming ItemId is in column 1
                        ProductName = worksheet.Cells[row, 2].Text,  // Assuming ProductName is in column 2
                        Location = worksheet.Cells[row, 4].Text,
                    };

                    var cellValue = worksheet.Cells[row, 3].Value;
                    int quantity = 0;

                    // Check if the cell value is a numeric type (int, double, etc.)
                    if (cellValue != null)
                    {
                        if (cellValue is int)
                        {
                            quantity = (int)cellValue;  // If the value is an integer
                        }
                        else if (cellValue is double || cellValue is float)
                        {
                            quantity = Convert.ToInt32(cellValue);  // Convert float/double to int
                        }
                        else
                        {
                            // Try parsing the string value if the type is not numeric
                            int.TryParse(cellValue.ToString(), out quantity);
                        }
                    }

                    item.Quantity = quantity; // Assign the parsed or default quantity
                    items.Add(item);
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any error that occurs while processing the file
            Console.WriteLine($"An error occurred while processing the file {filePath}: {ex.Message}");
            Console.WriteLine(ex.StackTrace);  // Optionally display the full stack trace
        }

        return items;
    }

    private int FindHeaderRow(ExcelWorksheet worksheet, int rowCount)
    {
        try
        {
            // Look for the first row with the expected headers
            for (int row = 1; row <= rowCount; row++)
            {
                // Use IndexOf for case-insensitive search
                if (
                    worksheet.Cells[row, 1].Text.IndexOf("ItemId", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    worksheet.Cells[row, 2].Text.IndexOf("ProductName", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    worksheet.Cells[row, 3].Text.IndexOf("Quantity", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    worksheet.Cells[row, 3].Text.IndexOf("Location", StringComparison.OrdinalIgnoreCase) >= 0

                    ) 
                {
                    return row;
                }
            }
            throw new Exception("Could not find the header row.");
        }
        catch (Exception ex)
        {
            // Handle any errors during the header search
            Console.WriteLine($"An error occurred while finding the header row: {ex.Message}");
            throw;  // Rethrow the exception so it can be handled higher up the call stack
        }
    }

    private Dictionary<string, int> GetColumnIndices(ExcelWorksheet worksheet, int headerRow, int columnCount)
    {
        //This method identifies the columns for ItemId, ProductName, and Quantity by comparing the header values with the expected names.
        var columnIndices = new Dictionary<string, int>();
        try
        {
            for (int col = 1; col <= columnCount; col++)
            {
                string header = worksheet.Cells[headerRow, col].Text.Trim().ToLower();
                if (header.Contains("itemid"))
                    columnIndices["ItemId"] = col;
                if (header.Contains("productname"))
                    columnIndices["ProductName"] = col;
                if (header.Contains("quantity"))
                    columnIndices["Quantity"] = col;
                if (header.Contains("location"))
                    columnIndices["Location"] = col;
            }

            // If any of the required columns are missing, add them with a default value
            if (!columnIndices.ContainsKey("ItemId")) columnIndices["ItemId"] = -1;
            if (!columnIndices.ContainsKey("ProductName")) columnIndices["ProductName"] = -1;
            if (!columnIndices.ContainsKey("Quantity")) columnIndices["Quantity"] = -1;
            if (!columnIndices.ContainsKey("Location")) columnIndices["Location"] = -1;
        }
        catch (Exception ex)
        {
            // Handle any error during column index identification
            Console.WriteLine($"An error occurred while getting column indices: {ex.Message}");
            throw;  // Rethrow the exception so it can be handled higher up the call stack
        }

        return columnIndices;
    }

    private void WriteConsolidatedData(List<InventoryItem> items, string outputFile)
    {
        try
        {
            // Create a new Excel package to write the consolidated data
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Consolidated Inventory");

                // Write the headers
                worksheet.Cells[1, 1].Value = "ItemId";
                worksheet.Cells[1, 2].Value = "ProductName";
                worksheet.Cells[1, 3].Value = "Quantity";
                worksheet.Cells[1, 4].Value = "Location";

                // Write the data
                for (int i = 0; i < items.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = items[i].ItemId;
                    worksheet.Cells[i + 2, 2].Value = items[i].ProductName;
                    worksheet.Cells[i + 2, 3].Value = items[i].Quantity;
                    worksheet.Cells[i + 2, 4].Value = items[i].Location;
                }

                // Save the package to the file
                FileInfo fileInfo = new FileInfo(outputFile);
                package.SaveAs(fileInfo);
            }
        }
        catch (Exception ex)
        {
            // Handle any error while saving the Excel file
            Console.WriteLine($"An error occurred while writing the consolidated data to the file: {ex.Message}");
            Console.WriteLine(ex.StackTrace);  // Optionally display the full stack trace
        }
    }
}

class InventoryItem
{
    public string ItemId { get; set; }
    public string ProductName { get; set; }
    public int Quantity { get; set; }
    public string  Location { get; set; }
}
