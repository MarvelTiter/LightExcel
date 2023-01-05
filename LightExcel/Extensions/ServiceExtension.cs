using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel
{
    public static class ServiceExtension
    {
        public static IServiceCollection AddLightExcel(this IServiceCollection services)
        {
            services.AddSingleton<IExcelHelper, ExcelHelper>();
            return services;
        }
    }
}
