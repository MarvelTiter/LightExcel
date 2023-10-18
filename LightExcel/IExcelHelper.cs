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
        // write
        void WriteExcel(string path, object data, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null);
        void WriteExcelByTemplate(string path, string template, object data, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null);
        ITransactionExcelHelper BeginTransaction(string path, Action<ExcelConfiguration>? config = null);

        // read
        IExcelDataReader ReadExcel(string path, string? sheetName = null, Action<ExcelConfiguration>? config = null);
        IEnumerable<T> QueryExcel<T>(string path, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null);
        IEnumerable<dynamic> QueryExcel(string path, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null);

        IExcelDataReader ReadExcel(Stream stream, string? sheetName = null, Action<ExcelConfiguration>? config = null);
        IEnumerable<T> QueryExcel<T>(Stream stream, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null);
        IEnumerable<dynamic> QueryExcel(Stream stream, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null);
    }
}
