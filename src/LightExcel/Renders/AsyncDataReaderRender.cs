using LightExcel.OpenXml;
using LightExcel.Utils;
using System.Data;
using System.Data.Common;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace LightExcel.Renders;
#if NET6_0_OR_GREATER
internal class AsyncDataReaderRender(ExcelConfiguration configuration) : AsyncRenderBase<DbDataReader, DbDataReader>(configuration)
{
    public override IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(DbDataReader data)
    {
        for (int i = 0; i < data.FieldCount; i++)
        {
            var name = data.GetName(i);
            var col = new ExcelColumnInfo(name);
            col.NumberFormat = Configuration.CheckCellNumberFormat(name);
            col.Type = data.GetFieldType(i);
            col.ColumnIndex = i + 1;
            AssignDynamicInfo(col);
            yield return col;
        }
    }

    public override async Task RenderAsync(DbDataReader datas, Sheet sheet, TransConfiguration configuration, CancellationToken cancellationToken)
    {
        if (sheet.Columns.Length == 0)
        {
            ExcelColumnInfo[] columns = [.. CollectExcelColumnInfo(datas)];
            sheet.Columns = columns;
        }

        var allRows = CollectALlRows(datas, sheet, sheet.Columns, configuration, cancellationToken);
        await sheet.WriteAsync(allRows, cancellationToken);
    }

    public override async IAsyncEnumerable<Row> RenderBodyAsync(DbDataReader datas, Sheet sheet, TransConfiguration configuration, [EnumeratorCancellation] CancellationToken cancellationToken)
    {
        var reader = datas;
        //var reader = data as IDataReader ?? throw new ArgumentException();
        var rowIndex = Configuration.StartRowIndex;
        var maxColumnIndex = 0;
        while (await reader.ReadAsync(cancellationToken))
        {
            var row = new Row() { RowIndex = ++rowIndex };
            var cellIndex = 0;
            foreach (var col in sheet.Columns)
            {
                if (col.Ignore) continue;
                var value = reader.GetValue(col.ColumnIndex - 1);
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

    private async IAsyncEnumerable<Row> CollectALlRows(DbDataReader reader, Sheet sheet, ExcelColumnInfo[] columns, TransConfiguration configuration, [EnumeratorCancellation] CancellationToken cancellationToken)
    {
        if (Configuration.UseHeader)
        {
            var headers = RenderHeader(columns, configuration);
            foreach (var row in headers)
            {
                yield return row;
            }
        }
        var rows = RenderBodyAsync(reader, sheet, configuration, cancellationToken);
        await foreach (var row in rows.WithCancellation(cancellationToken))
        {
            yield return row;
        }
    }
}
#endif