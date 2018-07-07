# abp-appfactory-excelexporter
ASP.NET Boilerplate module for exporting any IEnumerable to Excel.

## Installation
Add a reference in you application module to the Nuget package (https://www.nuget.org/packages/Abp.AppFactory.Interfaces)[Abp.AppFactory.Interfaces]


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