using Abp.AppFactory.BlobProvider.Storage;
using Abp.Application.Services.Dto;
using Abp.Dependency;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace App.Factory.ExcelExport
{
    public class Exporter<T> : ITransientDependency
    {
        private readonly BlobStorage blobStorage;

        public Exporter(BlobStorage blobStorage)
        {
            this.blobStorage = blobStorage;
        }

        public string[] Headings { private get; set; }

        public IEnumerable<T> Content { private get; set; }

        public async Task Export(string filename)
        {

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(filename);

                worksheet.Cells["A1"].LoadFromArrays(new List<string[]>() { Headings }).Style.Font.Bold = true;
                worksheet.Cells["A2"].LoadFromCollection(Content, false).AutoFitColumns();

                var byteArray = package.GetAsByteArray();
                var containerName = "imports";
                await blobStorage.UploadAsync(containerName, "excel", $"{filename}.xlsx", byteArray);

            }
        }


    }
}