using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel;

public static class ExcelHelperFactory
{
    private static readonly Lazy<IExcelHelper> _instance;
    static ExcelHelperFactory()
    {
        _instance = new Lazy<IExcelHelper>(() => new ExcelHelper());
    }

    public static IExcelHelper Instance => _instance.Value;
}
