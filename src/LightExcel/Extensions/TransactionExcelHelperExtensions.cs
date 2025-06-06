using LightExcel.Renders;
using LightExcel.Utils;
using System.Data;
#pragma warning disable IDE0130
namespace LightExcel;

public static partial class TransactionExcelHelperExtensions
{

    private static void WriteExcel<TRender>(ITransactionExcelHelper helper, object datas, string sheetName = "Sheet1", Action<TransConfiguration>? config = null)
        where TRender : IDataRender
    {
        var render = RenderCreator<TRender>.Create(helper.Configuration);
        var configuration = new TransConfiguration(helper.Configuration);
        config?.Invoke(configuration);
        helper.WriteExcel(render, datas, sheetName, configuration);
    }

    #region EnumerableEntityRender - 实体列表作为数据源

    /// <summary>
    /// 实体数据-保存到文件
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="helper"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcel<T>(this ITransactionExcelHelper helper, IEnumerable<T> datas, string sheetName = "Sheet1", Action<TransConfiguration>? config = null)
    {
        WriteExcel<EnumerableEntityRender<T>>(helper, datas, sheetName, config);
    }

    #endregion

    #region DataTableRender - DataTable作为数据源

    /// <summary>
    /// DataTable数据源-保存到文件
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcel(this ITransactionExcelHelper helper, DataTable datas, string sheetName = "Sheet1", Action<TransConfiguration>? config = null)
    {
        WriteExcel<DataTableRender>(helper, datas, sheetName, config);
    }

    #endregion

    #region DictionaryRender - Dictionary<string, object>作为数据源

    /// <summary>
    /// <![CDATA[Dictionary<string, object>]]>数据源
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcel(this ITransactionExcelHelper helper, IEnumerable<Dictionary<string, object>> datas, string sheetName = "Sheet1", Action<TransConfiguration>? config = null)
    {
        WriteExcel<DictionaryRender>(helper, datas, sheetName, config);
    }

    #endregion

    #region DataReaderRender - IDataReader作为数据源

    /// <summary>
    /// IDataReader数据源
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcel(this ITransactionExcelHelper helper, IDataReader datas, string sheetName = "Sheet1", Action<TransConfiguration>? config = null)
    {
        WriteExcel<DataReaderRender>(helper, datas, sheetName, config);
    }

    #endregion

#if NET6_0_OR_GREATER

    public static async Task WriteExcelAsync<T>(this ITransactionExcelHelper helper, IAsyncEnumerable<T> datas, string sheetName = "Sheet1", Action<TransConfiguration>? config = null, CancellationToken cancellationToken = default)
    {
        await helper.WriteExcelAsync<AsyncEnumerableEntityRender<T>>(datas, sheetName, config, cancellationToken);
    }

    public static async Task WriteExcelAsync(this ITransactionExcelHelper helper, IAsyncEnumerable<Dictionary<string, object?>> datas, string sheetName = "Sheet1", Action<TransConfiguration>? config = null, CancellationToken cancellationToken = default)
    {
        await helper.WriteExcelAsync<AsyncDictionaryRender>(datas, sheetName, config, cancellationToken);
    }

#endif
}

