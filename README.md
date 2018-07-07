# abp-appfactory-excelgenerator
ASP.NET Boilerplate module for generating and exporting any IEnumerable to Excel.

Wraps [EPPlus](https://github.com/JanKallman/EPPlus) into an ASP.NET Boilerplate module and saves the file using a [Blob Storage](https://azure.microsoft.com/en-us/services/storage/blobs/) provider module.

## Installation
Add a reference to the Nuget package [Abp.AppFactory.Interfaces](https://www.nuget.org/packages/Abp.AppFactory.Interfaces) where you want to inject the ExcelGenerator and inject the interface **IExcelGenerator**.

## Usage
```csharp
public DemoAppService: AsyncCrudService<DemoEntity, DemoEntityDto>
{
    private readonly IExcelGenerator excelGenerator;

    public DemoAppService(
        IBlobStorage blobStorage,
        IExcelGenerator excelGenerator

        IRepository<DemoEntity> repository,
    ) : base(repository)
    { 
        this.excelGenerator = demoGenerator;
        this.excelGenerator.SetStore(blobStorage);
    }

    public async Task<string> ExportAll()
    {
        var entities = Repository.GetAll();

        var dtos = ObjectMapper.Map<List<DemoEntityDto>>(entities);

        string url = await excelGenerator.CreateAndStoreAsync(dtos);

        return url;
    }
}
```