using LightExcel.OpenXml;
using LightExcel.Renders;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel;

public static partial class ExcelHeperExtensions
{
    #region DataReaderRender - IDataReader作为数据源

    /// <summary>
    /// IDataReader数据源
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcel(this IExcelHelper helper, string path, IDataReader datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcel<DataReaderRender>(helper, path, datas, sheetName, config);
    }

    /// <summary>
    /// IDataReader数据源
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcelByTemplate(this IExcelHelper helper, string path, string template, IDataReader datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcelByTemplate<DataReaderRender>(helper, path, template, datas, sheetName, config);
    }

    /// <summary>
    /// IDataReader数据源
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcel(this IExcelHelper helper, Stream stream, IDataReader datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcel<DataReaderRender>(helper, stream, datas, sheetName, config);
    }

    /// <summary>
    /// IDataReader数据源
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcelByTemplate(this IExcelHelper helper, Stream stream, Stream template, IDataReader datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcelByTemplate<DataReaderRender>(helper, stream, template, datas, sheetName, config);
    }

    /// <summary>
    /// IDataReader数据源
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcelByTemplate(this IExcelHelper helper, string path, Stream templateStream, IDataReader datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcelByTemplate<DataReaderRender>(helper, path, templateStream, datas, sheetName, config);
    }

    /// <summary>
    /// IDataReader数据源
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcelByTemplate(this IExcelHelper helper, Stream stream, string template, IDataReader datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcelByTemplate<DataReaderRender>(helper, stream, template, datas, sheetName, config);
    }
    #endregion

    #region 异步

#if NET6_0_OR_GREATER

    /// <summary>
    /// 实体数据-保存到文件
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static async Task WriteExcelAsync(this IExcelHelper helper, string path, DbDataReader datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null, CancellationToken cancellationToken = default)
    {
        using var trans = helper.BeginTransaction(path, config);
        //var render = new AsyncEnumerableEntityRender<T>(configuration);
        await trans.WriteExcelAsync<AsyncDataReaderRender>(datas, sheetName, cancellationToken: cancellationToken);

    }

    /// <summary>
    /// 实体数据-使用内存流
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="helper"></param>
    /// <param name="stream"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static async Task WriteExcelAsync(this IExcelHelper helper, Stream stream, DbDataReader datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null, CancellationToken cancellationToken = default)
    {
        //ExcelHelper.WriteExcel<EnumerableEntityRender<T>>(stream, datas, sheetName, config);
        using var trans = helper.BeginTransaction(stream, config);
        await trans.WriteExcelAsync<AsyncDataReaderRender>(datas, sheetName, cancellationToken: cancellationToken);
    }

    /// <summary>
    /// 实体数据-使用模板文件
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static async Task WriteExcelByTemplateAsync(this IExcelHelper helper, string path, string template, DbDataReader datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null, CancellationToken cancellationToken = default)
    {
        var configuration = new ExcelConfiguration();
        config?.Invoke(configuration);
        using var doc = ExcelDocument.CreateByTemplate(path, template, configuration);
        await TransExcelHelper.WriteByTemplateAsync<AsyncDataReaderRender>(doc, datas, sheetName, configuration, cancellationToken);
    }


#endif

    #endregion
}
