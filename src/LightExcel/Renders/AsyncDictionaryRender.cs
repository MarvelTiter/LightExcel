
using LightExcel.OpenXml;
using LightExcel.Utils;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

namespace LightExcel.Renders;
#if NET6_0_OR_GREATER
internal class AsyncDictionaryRender(ExcelConfiguration configuration) : AsyncEnumerableEntityRender<Dictionary<string, object?>>(configuration)
{
    public override IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(Dictionary<string, object?> data)
    {
        int index = 1;
        foreach (var item in data.Keys)
        {
            var col = new ExcelColumnInfo(item)
            {
                NumberFormat = Configuration.CheckCellNumberFormat(item),
                ColumnIndex = index++
            };
            AssignDynamicInfo(col);
            yield return col;
        }
    }

    public override async IAsyncEnumerable<Row> RenderBodyAsync(IAsyncEnumerable<Dictionary<string, object?>> data, Sheet sheet, TransConfiguration configuration, [EnumeratorCancellation]CancellationToken cancellationToken)
    {
        var rowIndex = Configuration.StartRowIndex;
        var maxColumnIndex = 0;
        await foreach (var item in data.WithCancellation(cancellationToken))
        {
            if (item is null) continue;
            var row = new Row() { RowIndex = ++rowIndex };
            var cellIndex = 0;
            foreach (var col in sheet.Columns)
            {
                if (col.Ignore) continue;
                item.TryGetValue(col.Name, out var value);
                cellIndex = col.ColumnIndex;
                var cell = CellHelper.CreateCell(cellIndex, rowIndex, value, col, configuration);
                row.AppendChild(cell);
            }
            maxColumnIndex = Math.Max(maxColumnIndex, cellIndex);
            yield return row;
        }
        sheet.MaxColumnIndex = maxColumnIndex;
        sheet.MaxRowIndex = rowIndex;
    }
}
#endif
