using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel
{
    public interface IExcelHelper
    {
        void WriteExcel(string path, object data, string sheetName = "sheet", bool appendSheet = true);
    }
}
