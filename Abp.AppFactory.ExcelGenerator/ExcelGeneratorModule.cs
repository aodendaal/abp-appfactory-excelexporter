using Abp.AppFactory.Interfaces;
using Abp.Modules;
using Abp.Reflection.Extensions;

namespace Abp.AppFactory.ExcelGenerator
{
    public class ExcelGeneratorModule : AbpModule
    {
        public override void PreInitialize()
        {
            IocManager.Register<IExcelGenerator, ExcelGenerator>(Dependency.DependencyLifeStyle.Transient);
        }

        public override void Initialize()
        {
            IocManager.RegisterAssemblyByConvention(typeof(ExcelGeneratorModule).GetAssembly());
        }
    }
}