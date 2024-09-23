using Azure.Identity;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;

public class Program
{
    static async Task Main(string[] args)
    {
        // Initialize the Graph client with interactive browser authentication
        var graphClient = GetGraphClient();

        // Get today's emails from the folder you are going into.
        // Note: The API cannot look into subfolders, so it needs to be a folder on the same level as those such as Inbox or Archive
        var results = await GetTodaysEmails(graphClient);

        // Output the results
        foreach (var result in results)
        {
            Console.WriteLine($"Email: {result.EmailName}, Time Sent: {result.TimeSent}, No Found: {result.NotFound}");
        }

        SaveResultsToExcel(results);
    }

    // Set up the Graph client with interactive browser authentication
    public static GraphServiceClient GetGraphClient()
    {
        var tenantId = "";  // Azure AD tenant ID
        var clientId = "";  // Azure AD application client ID

        var options = new InteractiveBrowserCredentialOptions
        {
            ClientId = clientId,
            TenantId = tenantId,
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            RedirectUri = new Uri("http://localhost")
        };

        var interactiveCredential = new InteractiveBrowserCredential(options);

        // Initialize GraphServiceClient
        return new GraphServiceClient(interactiveCredential);
    }

    // Fetch today's emails from the Desired folder
    public static async Task<List<EmailResult>> GetTodaysEmails(GraphServiceClient graphClient)
    {
        var results = new List<EmailResult>();

        // Get the desired mail folder
       var folderName = "";
        var mailFolders = await graphClient.Me.MailFolders.GetAsync();

        var CheckFolder = mailFolders.Value.FirstOrDefault(f => f.DisplayName == folderName);
        if (CheckFolder == null)
        {
            
            Console.WriteLine($"Folder '{folderName}' not found.");
            return results;
        }

        // Fetch messages from the desired folder
        var messages = await graphClient.Me.MailFolders[CheckFolder.Id].Messages
                                       .GetAsync((requestConfiguration) =>
                                       {
                                           requestConfiguration.QueryParameters.Top = 50;  // Retrieve top 50 messages
                                       });

        // Get today's date for filtering emails
        DateTime today = DateTime.UtcNow.Date;

        foreach (var message in messages.Value)
        {
            // Check if the email was received today
            if (message.ReceivedDateTime.HasValue && message.ReceivedDateTime.Value.Date == today)
            {
                Console.WriteLine($"Subject: {message.Subject}, Received: {message.ReceivedDateTime}");

                bool notFound = false;

                // Fetch attachments for each email
                var attachments = await graphClient.Me.Messages[message.Id].Attachments.GetAsync();

                // Check each attachment if it's a PDF
                foreach (var attachment in attachments.Value)
                {
                    if (attachment is FileAttachment fileAttachment && fileAttachment.Name.EndsWith(".pdf"))
                    {
                        byte[] fileBytes = fileAttachment.ContentBytes;
                        string tempPath = Path.Combine(Path.GetTempPath(), fileAttachment.Name);

                        File.WriteAllBytes(tempPath, fileBytes);

                        // Check if the PDF contains "Not Found"
                        notFound = CheckNotFound(tempPath);

                        File.Delete(tempPath);  // Clean up temporary files
                    }
                }

                // Add result to the list
                results.Add(new EmailResult
                {
                    EmailName = message.Subject,
                    TimeSent = message.ReceivedDateTime?.ToString(),
                    NotFound = notFound ? "Yes" : "No"
                });
            }
        }

        return results;
    }

    // Check if the PDF contains the "Not Found" text
    public static bool CheckNotFound(string pdfFilePath)
    {
        using (PdfReader reader = new PdfReader(pdfFilePath))
        using (PdfDocument pdfDoc = new PdfDocument(reader))
        {
            for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
            {
                string pageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(i));
                if (pageText.Contains("Not found"))
                {
                    return true;
                }
            }
        }
        return false;
    }

    public static void SaveResultsToExcel(List<EmailResult> results)
    {
        string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Results.xlsx");
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (ExcelPackage package = new ExcelPackage(new FileInfo("Test.xlsx")))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Results");
            worksheet.Cells[1, 1].Value = "Email Name";
            worksheet.Cells[1, 2].Value = "Time Sent";
            worksheet.Cells[1, 3].Value = "Not Found";

            int row = 2;
            foreach (var result in results)
            {
                worksheet.Cells[row, 1].Value = result.EmailName;
                worksheet.Cells[row, 2].Value = result.TimeSent;
                worksheet.Cells[row, 3].Value = result.NotFound;
                row++;
            }

            package.SaveAs(new FileInfo(excelPath));
            Console.WriteLine($"Results saved to {excelPath}");
        }
    }

}

// Class to store the email result
public class EmailResult
{
    public string EmailName { get; set; }
    public string TimeSent { get; set; }
    public string NotFound { get; set; }
}
