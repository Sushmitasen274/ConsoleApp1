using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Threading.Tasks;

class Program
{
    static async Task Main(string[] args)
    {
        //E:\ConsoleApp\FMS console Final by FMSA3\FMS console Final by FMSA3\ConsoleApp1\excel
        //string folderPath = @"E:\ConsoleApp\FMS console Final by FMSA3\FMS console Final by FMSA3\excelworkbook\files";
        //string outputFile = @"E:\ConsoleApp\FMS console Final by FMSA3\FMS console Final by FMSA3\excelworkbook\consolidated_inventory.xlsx";
        string inventoryPath = @"E:\ConsoleApp\FMS console Final by FMSA3\FMS console Final by FMSA3\ConsoleApp1\excel\InventoryFiles";
        string connectionString = "Server=192.168.6.16;Database=FMS_28042023;User Id=sa;Password=Idea@12345+;Encrypt=True;TrustServerCertificate=True;";

        // Step 1: Consolidate inventory data
        string[] files = Directory.GetFiles(inventoryPath, "*.xlsx");
        List<Task> inventoryTasks = new List<Task>();
        foreach (var file in files)
        {
            inventoryTasks.Add(Task.Run(() => ProcessInventoryFile(file, connectionString)));
        }
        await Task.WhenAll(inventoryTasks);
        Console.WriteLine("Inventory data consolidated.");


        var connectionStrings = new Dictionary<string, string>
 {
     { "HANA",  "Server=192.168.6.16;Database=FMS_28042023;User Id=sa;Password=Idea@12345+;Encrypt=True;TrustServerCertificate=True;" },
     { "Teradata",  "Server=192.168.6.16;Database=FMS_28042023;User Id=sa;Password=Idea@12345+;Encrypt=True;TrustServerCertificate=True;"},
     { "Snowflake", "Server=192.168.6.16;Database=FMS_28042023;User Id=sa;Password=Idea@12345+;Encrypt=True;TrustServerCertificate=True;"}
 };
        // Step 2: Consolidate sales data
        var salesDataTasks = new List<Task>
        {
            Task.Run(() => PullSalesDataFromHANA(connectionStrings["HANA"])),
            Task.Run(() => PullSalesDataFromTeradata(connectionStrings["Teradata"])),
            Task.Run(() => PullSalesDataFromSnowflake(connectionStrings["Snowflake"]))
        };
        await Task.WhenAll(salesDataTasks);
        Console.WriteLine("Sales data consolidated.");

        // Step 3: Tie out data and calculate variances
        CalculatePriceVariances(connectionString);

        // Step 4: Generate reports and send notifications
        GenerateAndSendReports(connectionString);

        Console.WriteLine("Process completed.");
    }

    static void ProcessInventoryFile(string filePath, string connectionString)
    {
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets[0];

        // Dynamically detect headers and data rows
        int headerRow = 1;
        int dataStartRow = 2; // Example: Adjust based on file

        var dataTable = new DataTable();
        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
        {
            dataTable.Columns.Add(worksheet.Cells[headerRow, col].Text);
        }

        for (int row = dataStartRow; row <= worksheet.Dimension.End.Row; row++)
        {
            var dataRow = dataTable.NewRow();
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                dataRow[col - 1] = worksheet.Cells[row, col].Text;
            }
            dataTable.Rows.Add(dataRow);
        }

        // Insert data into the database
        using var connection = new SqlConnection(connectionString);
        connection.Open();
        foreach (DataRow row in dataTable.Rows)
        {
            var command = new SqlCommand(
                "INSERT INTO InventoryData (ItemCode, ItemName, Quantity, FactoryCode, DateAdded) VALUES (@ItemCode, @ItemName, @Quantity, @FactoryCode, GETDATE())",
                connection);
            command.Parameters.AddWithValue("@ItemCode", row["ItemCode"]);
            command.Parameters.AddWithValue("@ItemName", row["ItemName"]);
            command.Parameters.AddWithValue("@Quantity", row["Quantity"]);
            command.Parameters.AddWithValue("@FactoryCode", row["FactoryCode"]);
            command.ExecuteNonQuery();
        }
    }

    static List<SalesRecord> PullSalesDataFromHANA(string connectionString)
    {
        // Example logic to pull data from HANA
        Console.WriteLine("Pulling data from HANA...");
       
            var records = new List<SalesRecord>();

            using (var connection = new Microsoft.Data.SqlClient.SqlConnection(connectionString))  // Use HANA-specific connection type
            {
                connection.Open();
                var command = new Microsoft.Data.SqlClient.SqlCommand("SELECT * FROM SalesData", connection);
                var reader = command.ExecuteReader();

                while (reader.Read())
                {
                    records.Add(new SalesRecord
                    {
                        ID = reader.GetInt32(0),        // SalesId from first column
                        SalesID = reader.GetInt32(1),      // ProductId from second column
                        FactoryCode = reader.GetString(2),   // ProductName from third column
                        MaterialCode = reader.GetDecimal(3),    // UnitPrice from fourth column
                        QuantitySold = reader.GetDecimal(4),  // TotalAmount from fifth column
                        SaleDate = reader.GetDateTime(5),    // SaleDate from sixth column
                        SalesPrice = reader.GetDecimal(6),  // CustomerName from seventh column
                      
                    });
                }
            }

            return records;
        }
    

    static void PullSalesDataFromTeradata(string connectionString)
    {
        // Example logic to pull data from Teradata
        Console.WriteLine("Pulling data from Teradata...");
        // Code to query Teradata and insert into SalesData table
    }

    static void PullSalesDataFromSnowflake(string connectionString)
    {
        // Example logic to pull data from Snowflake
        Console.WriteLine("Pulling data from Snowflake...");
        // Code to query Snowflake and insert into SalesData table
    }

    static void CalculatePriceVariances(string connectionString)
    {
        using var connection = new SqlConnection(connectionString);
        DbConnection.Open();

        string query = @"
            INSERT INTO PriceVariances (FactoryCode, MaterialCode, OriginalPrice, CurrentPrice, Variance, DateFlagged)
            SELECT
                mp.FactoryCode,
                mp.MaterialCode,
                mp.OriginalPrice,
                sd.SalesPrice,
                sd.SalesPrice - mp.OriginalPrice AS Variance,
                GETDATE()
            FROM MaterialPrices mp
            JOIN SalesData sd ON mp.MaterialCode = sd.MaterialCode AND mp.FactoryCode = sd.FactoryCode
            WHERE sd.SalesPrice <> mp.OriginalPrice;";

        using var command = new SqlCommand(query, connection);
        command.ExecuteNonQuery();
        Console.WriteLine("Price variances calculated.");
    }

    static void GenerateAndSendReports(string connectionString)
    {
        // Generate a consolidated report
        string reportPath = @"E:\ConsoleApp\FMS console Final by FMSA3\FMS console Final by FMSA3\ConsoleApp1\excel\MonthlyReport";
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("Report");

        using var connection = new SqlConnection(connectionString);
        connection.Open();
        string query = "SELECT * FROM PriceVariances";
        using var command = new SqlCommand(query, connection);
        using var reader = command.ExecuteReader();

        // Write data to Excel
        int row = 1;
        for (int col = 0; col < reader.FieldCount; col++)
        {
            worksheet.Cells[row, col + 1].Value = reader.GetName(col);
        }

        row++;
        while (reader.Read())
        {
            for (int col = 0; col < reader.FieldCount; col++)
            {
                worksheet.Cells[row, col + 1].Value = reader[col];
            }
            row++;
        }

        package.SaveAs(new FileInfo(reportPath));

        // Send the report via email
        //var toEmail = "shivkiranri@gmail.com";
        //var fromEmail = "shivkiran.rc@ideainfinityit.com";
        //var password = "Idea@123+";  // Make sure to replace 
        var mail = new MailMessage("shivkiranri@gmail.com", "shivkiran.rc@ideainfinityit.com")
        {
            Subject = "Monthly Price Variance Report",
            Body = "Please find the attached report."
        };
        mail.Attachments.Add(new Attachment(reportPath));

        using var smtp = new SmtpClient("mail.ieasybill.in")
        {
            Credentials = new System.Net.NetworkCredential("shivkiran.rc@ideainfinityit.com", "Idea@123+"),
            EnableSsl = true
        };
        smtp.Send(mail);
        Console.WriteLine("Report sent via email.");
    }
}
// Assuming you have a SalesRecord class to represent the data
public class SalesRecord
{
    public int ID { get; set; }
    public int SalesID { get; set; }
    public decimal SalesPrice { get; set; }
    public string FactoryCode { get; set; }
    public string MaterialCode { get; set; }
    public int QuantitySold { get; set; }
    public DateTime SaleDate { get; set; }
}

