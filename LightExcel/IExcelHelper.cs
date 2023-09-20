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
        void WriteExcel(object data, string? sheetName = null, Action<TransConfiguration>? config = null);
    }
    public interface IExcelHelper
    {
        void WriteExcel(string path, object data, string sheetName = "sheet", Action<ExcelConfiguration>? config = null);
        void WriteExcelByTemplate(string path, string template, object data, string sheetName = "sheet1", Action<ExcelConfiguration>? config = null);
        ITransactionExcelHelper BeginTransaction(string path, Action<ExcelConfiguration>? config = null);
        IExcelDataReader ReadExcel(string path, string? sheetName = null, Action<ExcelConfiguration>? config = null);
        IEnumerable<T> QueryExcel<T>(string path, string sheetName = "sheet1", Action<ExcelConfiguration>? config = null);
        IEnumerable<dynamic> QueryExcel(string path, string sheetName = "sheet1", Action<ExcelConfiguration>? config = null);
    }
}
