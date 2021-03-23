using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using Newtonsoft.Json;

namespace ExcelToOneDrive
{
    class Program
    {
        /// Excel handling:
        /// https://medium.com/swlh/openxml-sdk-brief-guide-c-7099e2391059

        /// Upload large files (>4MB) to onedrive
        /// https://medium.com/@cesarcodes/uploading-large-files-to-onedrive-with-microsoft-graph-api-33aeaa7a3319
        /// Requires following Graph application permissions: Files.ReadWrite.All is definitely needed (Sites.ReadWrite.All might also be needed)

        private static readonly string OneDrivePathAndFilename = "/UploadFolder/WorksheetName.xlsx";
        private static readonly string SheetName = "SheetName";

        private static readonly List<CustomData> CustomData = new List<CustomData>()
           {
               new CustomData() {Id="1001", Name="ABCD", City ="City1", Country="USA"},
               new CustomData() {Id="1002", Name="PQRS", City ="City2", Country="INDIA"},
               new CustomData() {Id="1003", Name="XYZZ", City ="City3", Country="CHINA"},
               new CustomData() {Id="1004", Name="LMNO", City ="City4", Country="UK"},
          };

        static async Task Main(string[] args)
        {
            #region Setup

            var env = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT");
            var builder = new ConfigurationBuilder()
                .AddJsonFile($"appsettings.json", true, true)
                .AddJsonFile($"appsettings.{env}.json", true, true)
                .AddUserSecrets<Program>()
                .AddEnvironmentVariables();

            var config = builder.Build();

            var confidentialClient = ConfidentialClientApplicationBuilder.CreateWithApplicationOptions(new ConfidentialClientApplicationOptions
            {
                ClientId = config["ClientId"],
                ClientSecret = config["ClientSecret"],
                TenantId = config["TenantId"]
            }).Build();

            var authProvider = new ClientCredentialProvider(confidentialClient);
            var graphServiceClient = new GraphServiceClient(authProvider);

            #endregion

            await UploadExcelForUser(graphServiceClient, config["Upn"]);
        }

        public static async Task UploadExcelForUser(GraphServiceClient graphServiceClient, string upn)
        {
            #region Create Excel in memory + Insert data

            #region Create Excel in memory

            MemoryStream ms = new MemoryStream();
            using var spreadsheetDocument = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);

            var workbookpart = spreadsheetDocument.AddWorkbookPart(); // Add a WorkbookPart to the document.
            workbookpart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
           
            var worksheetPart = workbookpart.AddNewPart<WorksheetPart>(); // Add a WorksheetPart to the WorkbookPart.
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);
            var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets()); // Add Sheets to the Workbook.
            var sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = SheetName
            };
            sheets.Append(sheet); // Append a new worksheet and associate it with the workbook.

            #endregion

            #region Insert data into Excel

            var table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(CustomData), typeof(DataTable));
            var headerRow = new Row();

            var columns = new List<string>();
            foreach (DataColumn column in table.Columns)
            {
                columns.Add(column.ColumnName);

                var cell = new Cell();
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(column.ColumnName);
                headerRow.AppendChild(cell);
            }

            sheetData.AppendChild(headerRow);

            foreach (DataRow dsrow in table.Rows)
            {
                var newRow = new Row();
                foreach (var col in columns)
                {
                    var cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(dsrow[col].ToString());
                    newRow.AppendChild(cell);
                }

                sheetData.AppendChild(newRow);
            }

            #endregion

            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();

            #endregion

            #region Upload to OneDrive

            var uploadSession = await GetUploadSession(graphServiceClient, OneDrivePathAndFilename, upn);

            var uploadProvider = new ChunkedUploadProvider(uploadSession, graphServiceClient, ms); // this is the upload manager class that does the magic
            var chunks = uploadProvider.GetUploadChunkRequests(); // these are the chunk requests that will be made
            var exceptions = new List<Exception>(); // you can use this to track exceptions, not used in this example

            foreach (var (chunk, i) in chunks.WithIndex()) // upload the chunks
            {
                var chunkRequestResponse = await uploadProvider.GetChunkRequestResponseAsync(chunk, exceptions);

                Console.WriteLine($"Uploading chunk {i} out of {chunks.Count()}");

                if (chunkRequestResponse.UploadSucceeded) // when the chunks are finished...
                {
                    Console.WriteLine("Upload is complete", chunkRequestResponse.ItemResponse);
                }
            }

            #endregion
        }

        public static async Task<UploadSession> GetUploadSession(GraphServiceClient client, string item, string upn)
        {
            return await client.Users[upn].Drive.Root.ItemWithPath(item).CreateUploadSession().Request().PostAsync();
        }
    }

    public static class Extensions
    {
        public static IEnumerable<(T item, int index)> WithIndex<T>(this IEnumerable<T> self)
        => self.Select((item, index) => (item, index));
    }
}
