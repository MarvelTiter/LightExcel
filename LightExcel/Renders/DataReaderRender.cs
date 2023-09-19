using LightExcel.OpenXml;
using LightExcel.Utils;
using System.Data;

namespace LightExcel.Renders
{
    internal class DataReaderRender : RenderBase
    {
        public override IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data, ExcelConfiguration configuration)
        {
            if (data is IDataReader d)
            {
                for (int i = 0; i < d.FieldCount; i++)
                {
                    var name = d.GetName(i);
                    var col = new ExcelColumnInfo(name);
                    col.NumberFormat = configuration.CheckCellNumberFormat(name);
                    col.Type = d.GetFieldType(i);
                    yield return col;
                }
            }
        }

        public override IEnumerable<Row> RenderBody(object data, Sheet sheet, IEnumerable<ExcelColumnInfo> columns, ExcelConfiguration configuration)
        {
            var reader = data as IDataReader ?? throw new ArgumentException();
            var rowIndex = configuration.StartRowIndex;
            var maxColumnIndex = 0;
            while (reader.Read())
            {
                var row = new Row() { RowIndex = ++rowIndex };
                var cellIndex = 1;
                foreach (var col in columns)
                {
                    if (col.Ignore) continue;
                    var cell = new Cell();
                    var value = reader.GetValue(reader.GetOrdinal(col.Name));
                    cellIndex = col.ColumnIndex ?? cellIndex;
                    cell.Reference = ReferenceHelper.ConvertXyToCellReference(cellIndex, rowIndex);
                    cell.Type = CellHelper.ConvertCellType(col.Type);
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