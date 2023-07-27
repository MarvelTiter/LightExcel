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
        void WriteExcel(string path, string template, object data);
        IExcelDataReader ReadExcel(string path);
        IEnumerable<T> QueryExcel<T>(string path, string? sheetName = null, int startRow = 2);
        IEnumerable<dynamic> QueryExcel(string path, string? sheetName = null, int startRow = 2);
    }
}
