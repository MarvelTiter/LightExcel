
using LightExcel.OpenXml;
using LightExcel.Utils;

namespace LightExcel.Renders
{
    internal class DictionaryRender : RenderBase
    {
        public override IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data, ExcelConfiguration configuration)
        {
            if (data is IEnumerable<Dictionary<string, object>> d)
            {
                foreach (var item in d.First().Keys)
                {
                    var col = new ExcelColumnInfo(item);
                    col.NumberFormat = configuration.CheckCellNumberFormat(item);
                    yield return col;
                }
            }
        }

        public override IEnumerable<Row> RenderBody(object data, Sheet sheet, IEnumerable<ExcelColumnInfo> columns, ExcelConfiguration configuration)
        {
            var values = data as IEnumerable<Dictionary<string, object>>;
            var rowIndex = configuration.StartRowIndex;
            var maxColumnIndex = 0;
            foreach (var item in values!)
            {
                if (item is null) continue;
                var row = new Row() { RowIndex = ++rowIndex };
                var cellIndex = 1;
                foreach (var col in columns)
                {
                    if (col.Ignore) continue;
                    var cell = new Cell();
                    var value = col.Name == "" ? "" : item[col.Name];
                    cellIndex = col.ColumnIndex ?? cellIndex;
                    cell.Reference = ReferenceHelper.ConvertXyToCellReference(cellIndex, rowIndex);
                    cell.Type = CellHelper.ConvertCellType(value?.GetType());
                    cell.Value = CellHelper.GetCellValue(col, value, configuration);
                    cell.StyleIndex = col.NumberFormat ? "1" : null;
                    row.AppendChild(cell);
                    cellIndex++;
                }
                maxColumnIndex = Math.Max(maxColumnIndex, cellIndex);
                yield return row;
            }
            sheet.MaxColumnIndex = maxColumnIndex;
            sheet.MaxRowIndex = rowIndex;
        }
    }
}
