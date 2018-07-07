using Abp.AppFactory.Interfaces;
using Abp.Dependency;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Abp.AppFactory.ExcelGenerator
{
    public class ExcelGenerator : IExcelGenerator, ITransientDependency
    {
        private const string defaultExportFilename = "export.xlsx";
        private const string defaultWorksheetName = "Sheet1";
        private readonly string[] defaultHeadings;

        private IBlobStorage defaultBlobStorage;
        private string defaultContainerName;
        private string defaultDirectory;

        public ExcelGenerator()
        {
            defaultHeadings = new string[] { };
        }

        #region SetStore

        public void SetStore(IBlobStorage blobStorage, string containerName, string directory = null)
        {
            this.defaultBlobStorage = blobStorage;
            this.defaultContainerName = containerName;
            this.defaultDirectory = directory;
        }

        #endregion SetStore

        #region CreateAsync

        public Task<byte[]> CreateAsync<T>(IEnumerable<T> content, string worksheetTitle = "Sheet1")
        {
            return CreateAsync(content, defaultHeadings, worksheetTitle);
        }

        public async Task<byte[]> CreateAsync<T>(IEnumerable<T> content, string[] headings, string worksheetTitle = "Sheet1")
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(worksheetTitle);

                if (headings.Length != 0)
                {
                    worksheet.Cells["A1"].LoadFromArrays(new List<string[]>() { headings }).Style.Font.Bold = true;
                    worksheet.Cells["A2"].LoadFromCollection(content, false).AutoFitColumns();
                }
                else
                {
                    worksheet.Cells["A1"].LoadFromCollection(content, true).AutoFitColumns();
                    worksheet.Cells["1:1"].Style.Font.Bold = true;
                }

                var byteArray = package.GetAsByteArray();

                return byteArray;
            }
        }

        #endregion CreateAsync

        #region CreateAndStoreAsync

        public Task<string> CreateAndStoreAsync<T>(IEnumerable<T> content, string filename = "export.xlsx", string worksheetTitle = "Sheet1")
        {
            return CreateAndStoreAsync(content, defaultHeadings, filename, worksheetTitle);
        }

        public Task<string> CreateAndStoreAsync<T>(IEnumerable<T> content, string[] headings, string filename = "export.xlsx", string worksheetTitle = "Sheet1")
        {
            return CreateAndStoreAsync(content, headings, defaultBlobStorage, defaultContainerName, defaultDirectory, filename, worksheetTitle);
        }

        public Task<string> CreateAndStoreAsync<T>(IEnumerable<T> content, IBlobStorage blobStorage, string containerName, string directory = null, string filename = "export.xlsx", string worksheetTitle = "Sheet1")
        {
            return CreateAndStoreAsync(content, defaultHeadings, blobStorage, containerName, directory, filename, worksheetTitle);
        }

        public async Task<string> CreateAndStoreAsync<T>(IEnumerable<T> content, string[] headings, IBlobStorage blobStorage, string containerName, string directory = null, string filename = "export.xlsx", string worksheetTitle = "Sheet1")
        {
            var bytes = await CreateAsync(content, headings, worksheetTitle);
            await StoreAsync(blobStorage, containerName, directory, filename, bytes);

            var url = Utilities.UrlPath.Combine(blobStorage.Endpoint, containerName);
            if (directory != null)
            {
                url = Utilities.UrlPath.Combine(url, directory);
            }
            url = Utilities.UrlPath.Combine(url, filename);

            return url;
        }

        #endregion CreateAndStore

        #region StoreAsync

        private async Task StoreAsync(IBlobStorage blobStorage, string containerName, string directory, string filename, byte[] bytes)
        {
            if (blobStorage == null)
            {
                throw new ArgumentNullException("blobStorage");
            }

            if (containerName == null)
            {
                throw new ArgumentNullException("containerName");
            }

            if (filename == null)
            {
                throw new ArgumentNullException("filename");
            }

            if (bytes == null)
            {
                throw new ArgumentNullException("bytes");
            }

            if (directory != null)
            {
                await blobStorage.UploadAsync(containerName, directory, filename, bytes);
            }
            else
            {
                await blobStorage.UploadAsync(containerName, filename, bytes);
            }
        }

        #endregion StoreAsync
    }
}