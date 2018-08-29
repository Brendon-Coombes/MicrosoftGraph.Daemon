using System;
using Microsoft.Extensions.Configuration;
using MicrosoftGraph.Daemon.Utilities;
using TERACC.WebJob.ResultsProcessor.Utilities;

namespace MicrosoftGraph.Daemon
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Simple console app to create a file and upload to a SharePoint Document Library");

            var configuration = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json")
                .Build();

            //Pull in configuration items
            string tenantId = configuration["TenantId"];
            string clientId = configuration["ClientId"];
            string clientSecret = configuration["ClientSecret"];
            string sharePointHostName = configuration["SharePointHostName"];
            string sharePointSiteRelativePath = configuration["SharePointSiteRelativePath"];
            string redirectUri = configuration["RedirectUri"];
            string authorityFormat = configuration["AuthorityFormat"];
            string graphScope = configuration["GraphScope"];
            string microsoftGraphBaseUrl = configuration["MicrosoftGraphBaseUrl"];

            //Create graph service and document generator
            IGraphService graphService = new GraphService(clientId, clientSecret, tenantId, redirectUri, authorityFormat, graphScope, microsoftGraphBaseUrl, sharePointHostName, sharePointSiteRelativePath);
            ExcelDocumentGenerator documentGenerator = new ExcelDocumentGenerator();

            //Generate a document and upload it to SharePoint
            var file = documentGenerator.GenerateDocument();
            var fileId = graphService.UploadFileAsync(file, "MyNewFile.xlsx", "{YOUR DOCUMENT LIBRARY NAME}").GetAwaiter().GetResult();

            Console.WriteLine("File written to SharePoint!");
            Console.WriteLine($"File Id: {fileId}");
            Console.ReadKey();
        }
    }
}
