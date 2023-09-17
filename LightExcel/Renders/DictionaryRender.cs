
using LightExcel.OpenXml;
using LightExcel.Utils;

namespace LightExcel.Renders
{
    internal class DictionaryRender : RenderBase, IDataRender
    {
        public void CollectExcelColumnInfo(object data, ExcelHelperConfiguration configuration)
        {
            if (data is IEnumerable<Dictionary<string, object>> d)
            {
                foreach (var item in d.First().Keys)
                {
                    Columns.Add(new ExcelColumnInfo(item));
                }
            }
        }

        public IEnumerable<Row> RenderBody(object data, Sheet sheet, ExcelHelperConfiguration configuration)
        {
            var values = data as IEnumerable<Dictionary<string, object>>;
            var rowIndex = configuration.UseHeader ? 1 : 0;
            var maxColumnIndex = 0;
            foreach (var item in values!)
            {
                if (item is null) continue;
                var row = new Row() { RowIndex = ++rowIndex };
                var cellIndex = 0;
                foreach (var col in Columns)
                {
                    if (col.Ignore) continue;
                    var cell = new Cell();
                    var value = item[col.Name];
                    cell.Reference = ReferenceHelper.ConvertXyToCellReference(++cellIndex, rowIndex);
                    cell.Type = CellHelper.ConvertCellType(value?.GetType());
                    cell.Value = CellHelper.GetCellValue(col, value, configuration);
                    row.AppendChild(cell);
                }
                maxColumnIndex = Math.Max(maxColumnIndex, cellIndex);
                yield return row;
            }
            sheet.MaxColumnIndex = maxColumnIndex;
            sheet.MaxRowIndex = rowIndex;
        }

        public Row RenderHeader(ExcelHelperConfiguration configuration)
        {
            var row = new Row() { RowIndex = 1 };
            var index = 0;
            foreach (var col in Columns)
            {
                var cell = new Cell
                {
                    Reference = ReferenceHelper.ConvertXyToCellReference(++index, 1),
                    Type = "str",
                    Value = col.Name
                };
                row.AppendChild(cell);
            }
            return row;
        }
    }
}
