using LightExcel.OpenXml;
using LightExcel.Renders;
using LightExcel.Utils;
using System.Data;
#pragma warning disable IDE0130
namespace LightExcel;
public static partial class ExcelHeperExtensions
{
    // 文件路径
    internal static void InternalWriteExcel<TRender>(IExcelHelper helper, string path, object datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
        where TRender : IDataRender
    {
        ExcelConfiguration configuration = new();
        config?.Invoke(configuration);
        var render = RenderCreator<TRender>.Create(configuration);
        helper.WriteExcel(render, path, datas, sheetName, configuration);
    }
    // 文件流
    internal static void InternalWriteExcel<TRender>(IExcelHelper helper, Stream stream, object datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
        where TRender : IDataRender
    {
        ExcelConfiguration configuration = new();
        config?.Invoke(configuration);
        var render = RenderCreator<TRender>.Create(configuration);
        helper.WriteExcel(render, stream, datas, sheetName, configuration);
    }
    // 文件路径+模板文件
    internal static void InternalWriteExcelByTemplate<TRender>(IExcelHelper helper, string path, string template, object datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
            where TRender : IDataRender
    {
        ExcelConfiguration configuration = new();
        config?.Invoke(configuration);
        var render = RenderCreator<TRender>.Create(configuration);
        using var doc = ExcelDocument.CreateByTemplate(path, template, configuration);
        helper.WriteExcelByTemplate(render, doc, datas, sheetName, configuration);
    }
    // 文件流+模板文件
    internal static void InternalWriteExcelByTemplate<TRender>(IExcelHelper helper, Stream stream, string template, object datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
        where TRender : IDataRender
    {
        ExcelConfiguration configuration = new();
        config?.Invoke(configuration);
        var render = RenderCreator<TRender>.Create(configuration);
        using var doc = ExcelDocument.CreateByTemplate(stream, template, configuration);
        helper.WriteExcelByTemplate(render, doc, datas, sheetName, configuration);
    }
    // 文件路径+模板文件流
    internal static void InternalWriteExcelByTemplate<TRender>(IExcelHelper helper, string path, Stream templateStream, object datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
        where TRender : IDataRender
    {
        ExcelConfiguration configuration = new();
        config?.Invoke(configuration);
        //using var doc = ExcelDocument.CreateByTemplate(stream, template, configuration);
        var render = RenderCreator<TRender>.Create(configuration);
        using var doc = ExcelDocument.CreateByTemplate(path, templateStream, configuration);
        helper.WriteExcelByTemplate(render, doc, datas, sheetName, configuration);
    }
    // 文件流+模板文件流
    internal static void InternalWriteExcelByTemplate<TRender>(IExcelHelper helper, Stream stream, Stream templateStream, object datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
        where TRender : IDataRender
    {
        ExcelConfiguration configuration = new();
        config?.Invoke(configuration);
        //using var doc = ExcelDocument.CreateByTemplate(stream, template, configuration);
        var render = RenderCreator<TRender>.Create(configuration);
        using var doc = ExcelDocument.CreateByTemplate(stream, templateStream, configuration);
        helper.WriteExcelByTemplate(render, doc, datas, sheetName, configuration);
    }


    #region BeginTransaction

    public static ITransactionExcelHelper BeginTransaction(this IExcelHelper helper, string path, Action<ExcelConfiguration>? config = null)
    {
        var configuration = new ExcelConfiguration();
        config?.Invoke(configuration);
        if (File.Exists(path)) File.Delete(path);
        var doc = ExcelDocument.Create(path, configuration);
        return helper.BeginTransaction(doc, configuration);
    }

    #endregion


#if NET6_0_OR_GREATER


    // TODO 模板重载

    

    //public static async Task WriteExcelByTemplateAsync(this IExcelHelper helper, string path, IAsyncEnumerable<Dictionary<string, object?>> data, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null, CancellationToken cancellationToken = default)
    //{
    //    using var trans = helper.BeginTransaction(path, config);
    //    //var render = new AsyncEnumerableEntityRender<T>(configuration);
    //    await trans.WriteExcelAsync<Dictionary<string, object?>, AsyncDictionaryRender>(data, sheetName, cancellationToken: cancellationToken);
    //}

#endif
}

