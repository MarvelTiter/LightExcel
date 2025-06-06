
using LightExcel.OpenXml;
using LightExcel.Utils;

namespace LightExcel.Renders
{
    internal class DictionaryRender(ExcelConfiguration configuration) : SyncRenderBase<IEnumerable<Dictionary<string, object?>>, Dictionary<string, object?>>(configuration)
    {
        public override IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(Dictionary<string, object?> data)
        {
            //if (data is not IEnumerable<Dictionary<string, object>> d)
            //{
            //    yield break;
            //}
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

        public override Dictionary<string, object?> GetFirstElement(IEnumerable<Dictionary<string, object?>> data)
            => data.First();

        public override IEnumerable<Row> RenderBody(IEnumerable<Dictionary<string, object?>> data, IRenderSheet sheet, TransConfiguration configuration)
        {
            //var values = data as IEnumerable<Dictionary<string, object>> ?? throw new ArgumentException();
            var values = data;
            var rowIndex = Configuration.StartRowIndex;
            var maxColumnIndex = 0;
            foreach (var item in values)
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
}
