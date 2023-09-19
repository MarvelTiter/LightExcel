using LightExcel.OpenXml;
using LightExcel.Utils;
using System.Data;

namespace LightExcel.Renders
{
    internal class DataTableRender : RenderBase
    {
        public override IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data, ExcelConfiguration configuration)
        {
            if (data is DataTable dt)
            {
                foreach (DataColumn column in dt.Columns)
                {
                    var col = new ExcelColumnInfo(column.ColumnName);
                    col.NumberFormat = configuration.CheckCellNumberFormat(column.ColumnName);
                    col.Type = column.DataType;
                    yield return col;
                }
            }
        }

        public override IEnumerable<Row> RenderBody(object data, Sheet sheet, IEnumerable<ExcelColumnInfo> columns, ExcelConfiguration configuration)
        {
            var values = data as DataTable;
            var rowIndex = configuration.StartRowIndex;
            var maxColumnIndex = 0;
            foreach (DataRow item in values!.Rows)
            {
                if (item is null) continue;
                var row = new Row() { RowIndex = ++rowIndex };
                var cellIndex = 1;
                foreach (var col in columns)
                {
                    if (col.Ignore) continue;
                    var cell = new Cell();
                    var value = item[col.Name];
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