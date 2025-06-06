using LightExcel.OpenXml;
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
        internal ExcelConfiguration Configuration { get; }
        void WriteExcel(IDataRender render, object data, string? sheetName = null, TransConfiguration? config = null);
#if NET6_0_OR_GREATER
        internal Task WriteExcelAsync<TRender>(object datas, string? sheetName = null, Action<TransConfiguration>? config = null, CancellationToken cancellationToken = default)
                where TRender : IAsyncDataRender;

        //internal Task WriteByTemplateAsync<T, TRender>(ExcelArchiveEntry doc, IAsyncEnumerable<T> data, string sheetName, ExcelConfiguration configuration, CancellationToken cancellationToken = default)
        //        where TRender : IAsyncDataRender<T>;
#endif
    }

    public interface IExcelHelper
    {
        #region 写入
        void WriteExcel(IDataRender render, string path, object data, string sheetName = "Sheet1", ExcelConfiguration? config = null);
        void WriteExcelByTemplate(IDataRender render, string path, string template, object data, string sheetName = "Sheet1", ExcelConfiguration? config = null);
        internal ITransactionExcelHelper BeginTransaction(ExcelArchiveEntry doc, ExcelConfiguration? config = null);
        // write stream
        void WriteExcel(IDataRender render, Stream stream, object data, string sheetName = "Sheet1", ExcelConfiguration? config = null);
        void WriteExcelByTemplate(IDataRender render, Stream stream, Stream templateStream, object data, string sheetName = "Sheet1", ExcelConfiguration? config = null);
        internal void WriteExcelByTemplate(IDataRender render, ExcelArchiveEntry doc, object data, string sheetName = "Sheet1", ExcelConfiguration? config = null);
        ITransactionExcelHelper BeginTransaction(Stream stream, Action<ExcelConfiguration>? config = null);
        #endregion

        #region 读取
        IExcelDataReader ReadExcel(string path, string? sheetName = null, Action<ExcelConfiguration>? config = null);
        IEnumerable<T> QueryExcel<T>(string path, string? sheetName, Action<ExcelConfiguration>? config = null);
        IEnumerable<dynamic> QueryExcel(string path, string? sheetName, Action<ExcelConfiguration>? config = null);

        IExcelDataReader ReadExcel(Stream stream, string? sheetName = null, Action<ExcelConfiguration>? config = null);
        IEnumerable<T> QueryExcel<T>(Stream stream, string? sheetName, Action<ExcelConfiguration>? config = null);
        IEnumerable<dynamic> QueryExcel(Stream stream, string? sheetName, Action<ExcelConfiguration>? config = null);
        #endregion
    }
}
