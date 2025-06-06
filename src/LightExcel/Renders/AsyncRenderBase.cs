using LightExcel.OpenXml;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace LightExcel.Renders;
#if NET6_0_OR_GREATER
internal abstract class AsyncRenderBase<TData, TElement>(ExcelConfiguration configuration) : RenderBase(configuration), IAsyncDataRender
{

    public virtual IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data)
        => CollectExcelColumnInfo((TElement)data);
    public virtual IAsyncEnumerable<Row> RenderBodyAsync(object datas, Sheet sheet, TransConfiguration configuration, CancellationToken cancellationToken)
        => RenderBodyAsync((TData)datas, sheet, configuration, cancellationToken);
    public virtual Task RenderAsync(object datas, Sheet sheet, TransConfiguration configuration, CancellationToken cancellationToken)
        => RenderAsync((TData)datas, sheet, configuration, cancellationToken);

    public abstract IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(TElement data);

    public abstract IAsyncEnumerable<Row> RenderBodyAsync(TData datas, Sheet sheet, TransConfiguration configuration, CancellationToken cancellationToken);

    public abstract Task RenderAsync(TData datas, Sheet sheet, TransConfiguration configuration, CancellationToken cancellationToken);
}

/// <summary>
/// 使用<see cref="IAsyncEnumerable<TElement>"/>作为数据源
/// </summary>
/// <typeparam name="TElement"></typeparam>
/// <param name="configuration"></param>
internal abstract class AsyncEnumerableRenderBase<TElement>(ExcelConfiguration configuration)
    : AsyncRenderBase<IAsyncEnumerable<TElement>, TElement>(configuration)
{
    public override async Task RenderAsync(IAsyncEnumerable<TElement> datas, Sheet sheet, TransConfiguration configuration, CancellationToken cancellationToken)
    {
        // 获取第一个元素（但不影响后续遍历）
        await using var enumerator = datas.GetAsyncEnumerator(cancellationToken);
        if (!await enumerator.MoveNextAsync())
        {
            return; // 无数据
        }
        var firstItem = enumerator.Current;
        if (sheet.Columns.Length == 0)
        {
            ExcelColumnInfo[] columns = [.. CollectExcelColumnInfo(firstItem)];
            sheet.Columns = columns;
        }

        var allRows = CollectALlRows(GetRemainingData(enumerator), sheet, sheet.Columns, configuration, cancellationToken);
        await sheet.WriteAsync(allRows, cancellationToken);
    }

    private async IAsyncEnumerable<Row> CollectALlRows(IAsyncEnumerable<TElement> datas, Sheet sheet, ExcelColumnInfo[] columns, TransConfiguration configuration, [EnumeratorCancellation] CancellationToken cancellationToken)
    {
        if (Configuration.UseHeader)
        {
            var headers = RenderHeader(columns, configuration);
            foreach (var row in headers)
            {
                yield return row;
            }
        }
        var rows = RenderBodyAsync(datas, sheet, configuration, cancellationToken);
        await foreach (var row in rows.WithCancellation(cancellationToken))
        {
            yield return row;
        }
    }

    private static async IAsyncEnumerable<TElement> GetRemainingData(IAsyncEnumerator<TElement> enumerator)
    {
        // 已经 MoveNext 一次，所以直接 yield 剩余数据
        do
        {
            yield return enumerator.Current;
        } while (await enumerator.MoveNextAsync());
    }
}

#endif
