using LightExcel.OpenXml;
using LightExcel.Utils;
using System.Data;

namespace LightExcel.Renders
{
    internal class DataReaderRender(ExcelConfiguration configuration) : SyncRenderBase<IDataReader, IDataReader>(configuration)
    {
        public override IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(IDataReader data)
        {
            //if (data is not IDataReader d)
            //{
            //    yield break;
            //}
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
        public override IDataReader GetFirstElement(IDataReader data) => data;

        public override IEnumerable<Row> RenderBody(IDataReader data, IRenderSheet sheet, TransConfiguration configuration)
        {
            //var reader = data as IDataReader ?? throw new ArgumentException();
            var reader = data;
            var rowIndex = Configuration.StartRowIndex;
            var maxColumnIndex = 0;
            while (reader.Read())
            {
                var row = new Row() { RowIndex = ++rowIndex };
                var cellIndex = 0;
                foreach (var col in sheet.Columns)
                {
                    if (col.Ignore) continue;
                    var value = reader.GetValue(col.ColumnIndex - 1);
                    cellIndex = col.ColumnIndex;
                    var cell = CellHelper.CreateCell(cellIndex, rowIndex, value, col, configuration);
                    //var cell = new Cell();
                    //var (v, t) = CellHelper.FormatCell(value, Configuration, col);
                    ////cell.Type = CellHelper.ConvertCellType(col.Type);
                    ////cell.Value = CellHelper.GetCellValue(col, value, Configuration);
                    //cell.Value = v;
                    //cell.Type = t;
                    //cell.StyleIndex = col.NumberFormat || configuration.NumberFormatColumnFilter(col) ? "1" : null;
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