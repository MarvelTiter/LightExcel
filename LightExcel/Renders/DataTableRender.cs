using LightExcel.OpenXml;
using LightExcel.Utils;
using System.Data;

namespace LightExcel.Renders
{
    internal class DataTableRender : RenderBase
    {
        public DataTableRender(ExcelConfiguration configuration) : base(configuration) { }
        public override IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data)
        {
            if (data is DataTable dt)
            {
                int index = 1;
                foreach (DataColumn column in dt.Columns)
                {
                    var col = new ExcelColumnInfo(column.ColumnName);
                    col.NumberFormat = Configuration.CheckCellNumberFormat(column.ColumnName, dt.TableName);
                    col.Type = column.DataType;
                    col.ColumnIndex = index++;
                    yield return col;
                }
            }
        }

        public override IEnumerable<Row> RenderBody(object data, Sheet sheet, IEnumerable<ExcelColumnInfo> columns, TransConfiguration configuration)
        {
            var values = data as DataTable;
            var rowIndex = Configuration.StartRowIndex;
            var maxColumnIndex = 0;
            foreach (DataRow item in values!.Rows)
            {
                if (item is null) continue;
                var row = new Row() { RowIndex = ++rowIndex };
                var cellIndex = 0;
                foreach (var col in columns)
                {
                    if (col.Ignore) continue;
                    var value = item[col.Name];
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