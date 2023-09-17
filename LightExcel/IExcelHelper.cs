using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel
{
    public interface ITransactionExcelHelper : IDisposable
    {
        void WriteExcel(object data, string? sheetName = null);
    }
    public interface IExcelHelper
    {
        void WriteExcel(string path, object data, string sheetName = "sheet1", Action<ExcelHelperConfiguration>? config = null);
        void WriteExcelByTemplate(string path, string template, object data, string sheetName = "sheet1", Action<ExcelHelperConfiguration>? config = null);
        ITransactionExcelHelper BeginTransaction(string path, Action<ExcelHelperConfiguration>? config = null);
        IExcelDataReader ReadExcel(string path);
        IEnumerable<T> QueryExcel<T>(string path, string? sheetName = null, int startRow = 2);
        IEnumerable<dynamic> QueryExcel(string path, string? sheetName = null, int startRow = 2);
    }
}
