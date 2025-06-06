using LightExcel.OpenXml;
using LightExcel.Renders;
using LightExcel.Utils;
using System.Data;
using System.IO;

namespace LightExcel;

public static partial class ExcelHeperExtensions
{
    #region EnumerableEntityRender - 实体列表作为数据源

    /// <summary>
    /// 实体数据-保存到文件
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcel<T>(this IExcelHelper helper, string path, IEnumerable<T> datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcel<EnumerableEntityRender<T>>(helper, path, datas, sheetName, config);
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
    public static void WriteExcel<T>(this IExcelHelper helper, Stream stream, IEnumerable<T> datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        //ExcelHelper.WriteExcel<EnumerableEntityRender<T>>(stream, datas, sheetName, config);
        InternalWriteExcel<EnumerableEntityRender<T>>(helper, stream, datas, sheetName, config);
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
    public static void WriteExcelByTemplate<T>(this IExcelHelper helper, string path, string template, IEnumerable<T> datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcelByTemplate<EnumerableEntityRender<T>>(helper, path, template, datas, sheetName, config);
    }

    /// <summary>
    /// 实体数据-写入和模板都使用内存流
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcelByTemplate<T>(this IExcelHelper helper, Stream stream, Stream templateStream, IEnumerable<T> datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcelByTemplate<EnumerableEntityRender<T>>(helper, stream, templateStream, datas, sheetName, config);
    }

    /// <summary>
    /// 实体数据-模板使用流，保存到文件
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcelByTemplate<T>(this IExcelHelper helper, string path, Stream templateStream, IEnumerable<T> datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcelByTemplate<EnumerableEntityRender<T>>(helper, path, templateStream, datas, sheetName, config);
    }

    /// <summary>
    /// 实体数据-使用模板文件，保存到内存流
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="helper"></param>
    /// <param name="stream"></param>
    /// <param name="template"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcelByTemplate<T>(this IExcelHelper helper, Stream stream, string template, IEnumerable<T> datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcelByTemplate<EnumerableEntityRender<T>>(helper, stream, template, datas, sheetName, config);
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
    public static async Task WriteExcelAsync<T>(this IExcelHelper helper, string path, IAsyncEnumerable<T> datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null, CancellationToken cancellationToken = default)
    {
        using var trans = helper.BeginTransaction(path, config);
        //var render = new AsyncEnumerableEntityRender<T>(configuration);
        await trans.WriteExcelAsync<AsyncEnumerableEntityRender<T>>(datas, sheetName, cancellationToken: cancellationToken);

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
    public static async Task WriteExcelAsync<T>(this IExcelHelper helper, Stream stream, IAsyncEnumerable<T> datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null, CancellationToken cancellationToken = default)
    {
        //ExcelHelper.WriteExcel<EnumerableEntityRender<T>>(stream, datas, sheetName, config);
        using var trans = helper.BeginTransaction(stream, config);
        await trans.WriteExcelAsync<AsyncEnumerableEntityRender<T>>(datas, sheetName, cancellationToken: cancellationToken);
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
    public static async Task WriteExcelByTemplateAsync<T>(this IExcelHelper helper, string path, string template, IAsyncEnumerable<T> datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null, CancellationToken cancellationToken = default)
    {
        var configuration = new ExcelConfiguration();
        config?.Invoke(configuration);
        using var doc = ExcelDocument.CreateByTemplate(path, template, configuration);
        await TransExcelHelper.WriteByTemplateAsync<AsyncEnumerableEntityRender<T>>(doc, datas, sheetName, configuration, cancellationToken);
    }


#endif

    #endregion
}
