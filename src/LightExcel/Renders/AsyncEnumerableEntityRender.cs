using LightExcel.Attributes;
using LightExcel.OpenXml;
using LightExcel.Utils;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

#if NET6_0_OR_GREATER
namespace LightExcel.Renders
{
    internal class AsyncEnumerableEntityRender<T>(ExcelConfiguration configuration) : AsyncEnumerableRenderBase<T>(configuration)
    {
        private readonly Type elementType = typeof(T);

        public override IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(T data)
        {
            var properties = elementType.GetProperties();
            int index = 1;
            foreach (var prop in properties)
            {
                ExcelColumnAttribute? excelColumnAttribute = prop.GetCustomAttribute<ExcelColumnAttribute>();
                if (excelColumnAttribute?.Ignore ?? false) continue;
#if NET6_0_OR_GREATER
                var displayAttribute = prop.GetCustomAttribute<System.ComponentModel.DataAnnotations.DisplayAttribute>();
                var col = new ExcelColumnInfo(excelColumnAttribute?.Name ?? displayAttribute?.Name ?? prop.Name);
#else
                var col = new ExcelColumnInfo(excelColumnAttribute?.Name ??  prop.Name);
#endif

                col.Ignore = excelColumnAttribute?.Ignore ?? false;
                col.Property = new Property(prop);
                col.Type = prop.PropertyType;
                col.NumberFormat = excelColumnAttribute?.NumberFormat ?? false;
                col.Format = excelColumnAttribute?.Format;
                col.ColumnIndex = index++;
                col.AutoWidth = excelColumnAttribute?.AutoWidth ?? false;
                col.Width = excelColumnAttribute?.Width;
                AssignDynamicInfo(col);
                yield return col;
            }
        }


        public override async IAsyncEnumerable<Row> RenderBodyAsync(IAsyncEnumerable<T> data, Sheet sheet, TransConfiguration configuration, [EnumeratorCancellation] CancellationToken cancellationToken)
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
                    if (col.Property == null)
                    {
                        var p = elementType.GetProperty(col.Name);
                        if (p == null) continue;
                        col.Property = new Property(p);
                    }
                    var value = col.Property.GetValue(item);
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
}
#endif