using Abp.AppFactory.BlobProvider;
using Abp.Modules;
using Abp.Reflection.Extensions;
using App.Factory.ExcelExport;

namespace Abp.AppFactory.ExcelExport
{
    [DependsOn(typeof(BlobProviderModule))]
    public class ExcelExporterModule : AbpModule
    {
        public override void Initialize()
        {
            IocManager.Register(typeof(Exporter<>));
            IocManager.RegisterAssemblyByConvention(typeof(ExcelExporterModule).GetAssembly());
        }
    }
}