// See https://aka.ms/new-console-template for more information
using System.Net;
using System.Net.Mail;
using OfficeOpenXml;
using System;
using System.IO;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

public class Program : EmailSender
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
            EmailSender.SendEmailWithAttachment(outputFile);
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
            // Get all the .xlsx files in the specified folder
            var files = Directory.GetFiles(folderPath, "*.xlsx");

            // Check if no files were found
            if (files.Length == 0)
            {
                Console.WriteLine("No .xlsx files found in the specified folder.");
                return;  // Exit the method early as there's no data to process
            }

            var tasks = new List<Task<List<InventoryItem>>>();

            // Process each file concurrently
            foreach (var file in files)
            {
                tasks.Add(Task.Run(() => ProcessFile(file)));
            }

            // Await all tasks to complete
            var results = await Task.WhenAll(tasks);

            // If the results list is empty, display a message and return
            if (results.All(result => result == null || !result.Any()))
            {
                Console.WriteLine("No inventory data found in the provided files.");
                return;  // Exit the method as there's no valid data to write
            }

            // Consolidate all results into one list
            var allItems = results.SelectMany(result => result).ToList();

            // Write the consolidated data to a new file
            WriteConsolidatedData(allItems, outputFile);

            Console.WriteLine("Consolidation complete!");
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
            // Check if file exists before processing
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return items;
            }

            // Read the Excel file using EPPlus
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Ensure the workbook contains at least one worksheet
                var worksheets = package.Workbook.Worksheets;
                if (worksheets.Count == 0)
                {
                    Console.WriteLine($"The file {filePath} does not contain any worksheets.");
                    return items;  // Return early as there's nothing to process
                }

                // Iterate through the worksheets or select the first one dynamically
                foreach (var worksheet in worksheets)
                {
                    Console.WriteLine($"Processing worksheet: {worksheet.Name}");

                    int rowCount = worksheet.Dimension.Rows;
                    int columnCount = worksheet.Dimension.Columns;

                    // Step 1: Find the header row (you may need to adjust this depending on file structure)
                    int headerRow = FindHeaderRow(worksheet, rowCount);
                    if (headerRow == -1)  // If header row is not found
                    {
                        Console.WriteLine($"No valid header row found in sheet: {worksheet.Name} in file: {filePath}");
                        continue;  // Skip this worksheet
                    }

                    // Step 2: Identify relevant columns by header names
                    var columnIndices = GetColumnIndices(worksheet, headerRow, columnCount);

                    // Step 3: Process each data row
                    for (int row = headerRow + 1; row <= rowCount; row++)  // Skip header row
                    {
                        // Ensure the row contains valid data (check if any required column is empty)
                        var item = new InventoryItem
                        {
                            Location = GetCellValue(worksheet, row, columnIndices["Location"]),  // Replacing ItemId with Location
                            ProductName = GetCellValue(worksheet, row, columnIndices["ProductName"]),
                        };

                        // Skip rows with missing essential data
                        if (string.IsNullOrEmpty(item.Location) || string.IsNullOrEmpty(item.ProductName))
                        {
                            Console.WriteLine($"Skipping row {row} due to missing essential data in sheet: {worksheet.Name} in file: {filePath}");
                            continue;
                        }

                        var cellValue = worksheet.Cells[row, columnIndices["Quantity"]].Value;
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
        }
        catch (FileNotFoundException fnfEx)
        {
            // Specific exception handling for file not found
            Console.WriteLine($"File not found: {filePath}. Error: {fnfEx.Message}");
        }
        catch (IOException ioEx)
        {
            // Specific exception handling for IO errors
            Console.WriteLine($"Error reading the file {filePath}: {ioEx.Message}");
        }
        catch (Exception ex)
        {
            // Handle any other errors that occur while processing the file
            Console.WriteLine($"An error occurred while processing the file {filePath}: {ex.Message}");
            Console.WriteLine(ex.StackTrace);  // Optionally display the full stack trace
        }

        return items;
    }



    private string GetCellValue(ExcelWorksheet worksheet, int row, int column)
    {
        // If the column index is invalid (e.g., -1 for missing columns), return an empty string
        if (column == -1)
        {
            return string.Empty;
        }

        // Get the cell value as a string, trimming any leading/trailing whitespace
        return worksheet.Cells[row, column].Text.Trim();
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
                    worksheet.Cells[row, 1].Text.IndexOf("Location", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    worksheet.Cells[row, 2].Text.IndexOf("ProductName", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    worksheet.Cells[row, 3].Text.IndexOf("Quantity", StringComparison.OrdinalIgnoreCase) >= 0)
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
        //This method identifies the columns for Location, ProductName, and Quantity by comparing the header values with the expected names.
        var columnIndices = new Dictionary<string, int>();
        try
        {
            for (int col = 1; col <= columnCount; col++)
            {
                string header = worksheet.Cells[headerRow, col].Text.Trim().ToLower();
                if (header.Contains("location"))
                    columnIndices["Location"] = col;
                if (header.Contains("productname"))
                    columnIndices["ProductName"] = col;
                if (header.Contains("quantity"))
                    columnIndices["Quantity"] = col;
            }

            // If any of the required columns are missing, add them with a default value
            if (!columnIndices.ContainsKey("Location")) columnIndices["Location"] = -1;
            if (!columnIndices.ContainsKey("ProductName")) columnIndices["ProductName"] = -1;
            if (!columnIndices.ContainsKey("Quantity")) columnIndices["Quantity"] = -1;
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
                worksheet.Cells[1, 1].Value = "Location";
                worksheet.Cells[1, 2].Value = "ProductName";
                worksheet.Cells[1, 3].Value = "Quantity";

                // Write the data
                for (int i = 0; i < items.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = items[i].Location;
                    worksheet.Cells[i + 2, 2].Value = items[i].ProductName;
                    worksheet.Cells[i + 2, 3].Value = items[i].Quantity;
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
    public string? Location { get; set; }
    public string? ProductName { get; set; }
    public int Quantity { get; set; }

}


    public class EmailSender
    {
        public static void SendEmailWithAttachment(string attachmentPath)
        {
            var toEmail = "shivkiranri@gmail.com";
            var fromEmail = "shivkiran.rc@ideainfinityit.com";
            var password = "Idea@123+";  // Make sure to replace with actual password (or use App password for Gmail)

            try
            {
                // Set up the SMTP client using your provided email credentials
                var smtpClient = new SmtpClient("mail.ieasybill.in")
                {
                    Port = 587, // TLS port (use 465 for SSL if necessary)
                    Credentials = new NetworkCredential(fromEmail, password),  // From email and password for authentication
                    EnableSsl = true,  // Enable SSL/TLS encryption for security
                    Timeout = 5000  // Optional: Set a timeout for connection attempts (5 seconds)
                };

                // Create the mail message
                var mailMessage = new MailMessage
                {
                    From = new MailAddress(fromEmail),
                    Subject = "Consolidate Report by SushmitaSen",  // Subject of the email
                    Body = "Please find the attached report.",  // Body of the email
                    IsBodyHtml = true  // Email body as HTML (set to false for plain text)
                };

                // Add the recipient email address
                mailMessage.To.Add(toEmail);

                // Attach the generated Excel report
                if (File.Exists(attachmentPath))
                {
                    mailMessage.Attachments.Add(new Attachment(attachmentPath));  // Attach the file
                }
                else
                {
                    Console.WriteLine("Attachment not found.");
                }

                // Send the email
                smtpClient.Send(mailMessage);
                Console.WriteLine("Email sent successfully to " + toEmail);
            }
            catch (SmtpException smtpEx)
            {
                Console.WriteLine("SMTP error occurred: " + smtpEx.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while sending the email: " + ex.Message);
            }
        }
    }



