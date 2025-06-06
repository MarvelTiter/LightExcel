using LightExcel.OpenXml;
using LightExcel.Renders;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel;

public static partial class ExcelHeperExtensions
{
    #region DictionaryRender - Dictionary<string, object>作为数据源

    /// <summary>
    /// <![CDATA[Dictionary<string, object>]]>数据源
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcel(this IExcelHelper helper, string path, IEnumerable<Dictionary<string, object>> datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcel<DictionaryRender>(helper, path, datas, sheetName, config);
    }

    /// <summary>
    /// <![CDATA[Dictionary<string, object>]]>数据源
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcelByTemplate(this IExcelHelper helper, string path, string template, IEnumerable<Dictionary<string, object>> datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcelByTemplate<DictionaryRender>(helper, path, template, datas, sheetName, config);
    }

    /// <summary>
    /// <![CDATA[Dictionary<string, object>]]>数据源
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcel(this IExcelHelper helper, Stream stream, IEnumerable<Dictionary<string, object>> datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcel<DictionaryRender>(helper, stream, datas, sheetName, config);
    }

    /// <summary>
    /// <![CDATA[Dictionary<string, object>]]>数据源
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcelByTemplate(this IExcelHelper helper, Stream stream, Stream template, IEnumerable<Dictionary<string, object>> datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcelByTemplate<DictionaryRender>(helper, stream, template, datas, sheetName, config);
    }

    /// <summary>
    /// <![CDATA[Dictionary<string, object>]]>数据源
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcelByTemplate(this IExcelHelper helper, string path, Stream templateStream, IEnumerable<Dictionary<string, object>> datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcelByTemplate<DictionaryRender>(helper, path, templateStream, datas, sheetName, config);
    }

    /// <summary>
    /// <![CDATA[Dictionary<string, object>]]>数据源
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcelByTemplate(this IExcelHelper helper, Stream stream, string template, IEnumerable<Dictionary<string, object>> datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcelByTemplate<DictionaryRender>(helper, stream, template, datas, sheetName, config);
    }

    #endregion

    #region 异步

#if NET6_0_OR_GREATER

    public static async Task WriteExcelAsync(this IExcelHelper helper, string path, IAsyncEnumerable<Dictionary<string, object?>> data, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null, CancellationToken cancellationToken = default)
    {
        using var trans = helper.BeginTransaction(path, config);
        //var render = new AsyncEnumerableEntityRender<T>(configuration);
        await trans.WriteExcelAsync<AsyncDictionaryRender>(data, sheetName, cancellationToken: cancellationToken);
    }

    public static async Task WriteExcelAsync(this IExcelHelper helper, Stream stream, IAsyncEnumerable<Dictionary<string, object?>> data, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null, CancellationToken cancellationToken = default)
    {
        using var trans = helper.BeginTransaction(stream, config);
        //var render = new AsyncEnumerableEntityRender<T>(configuration);
        await trans.WriteExcelAsync<AsyncDictionaryRender>(data, sheetName, cancellationToken: cancellationToken);
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
    public static async Task WriteExcelByTemplateAsync(this IExcelHelper helper, string path, string template, IAsyncEnumerable<Dictionary<string, object?>> datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null, CancellationToken cancellationToken = default)
    {
        var configuration = new ExcelConfiguration();
        config?.Invoke(configuration);
        using var doc = ExcelDocument.CreateByTemplate(path, template, configuration);
        await TransExcelHelper.WriteByTemplateAsync<AsyncDictionaryRender>(doc, datas, sheetName, configuration, cancellationToken);
    }
#endif

    #endregion
}
