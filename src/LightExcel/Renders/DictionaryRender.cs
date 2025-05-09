
using LightExcel.OpenXml;
using LightExcel.Utils;

namespace LightExcel.Renders
{
    internal class DictionaryRender : RenderBase
    {
        public DictionaryRender(ExcelConfiguration configuration) : base(configuration) { }

        public override IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data)
        {
            if (data is IEnumerable<Dictionary<string, object>> d)
            {
                int index = 1;
                foreach (var item in d.First().Keys)
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
        }

        public override IEnumerable<Row> RenderBody(object data, Sheet sheet, ExcelColumnInfo[] columns, TransConfiguration configuration)
        {
            var values = data as IEnumerable<Dictionary<string, object>> ?? throw new ArgumentException();
            var rowIndex = Configuration.StartRowIndex;
            var maxColumnIndex = 0;
            foreach (var item in values!)
            {
                if (item is null) continue;
                var row = new Row() { RowIndex = ++rowIndex };
                var cellIndex = 0;
                foreach (var col in columns)
                {
                    if (col.Ignore) continue;
                    item.TryGetValue(col.Name, out var value);
                    cellIndex = col.ColumnIndex;
                    var nf = configuration.NumberFormatColumnFilter(col);
                    var cell = CellHelper.CreateCell(cellIndex, rowIndex, value, col, nf, Configuration);
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
