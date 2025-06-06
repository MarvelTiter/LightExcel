using LightExcel.OpenXml;
using LightExcel.Utils;
using System.Data;

namespace LightExcel.Renders
{
    internal class DataTableRender(ExcelConfiguration configuration) : SyncRenderBase<DataTable, DataRow>(configuration)
    {
        public override IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(DataRow data)
        {
            //if (data is not DataTable dt)
            //{
            //    yield break;
            //}
            var dt = data.Table;
            int index = 1;
            foreach (DataColumn column in dt.Columns)
            {
                var col = new ExcelColumnInfo(column.ColumnName);
                col.NumberFormat = Configuration.CheckCellNumberFormat(column.ColumnName, dt.TableName);
                col.Type = column.DataType;
                col.ColumnIndex = index++;
                AssignDynamicInfo(col);
                yield return col;
            }
        }

        public override DataRow GetFirstElement(DataTable data) => data.NewRow();

        public override IEnumerable<Row> RenderBody(DataTable data, IRenderSheet sheet, TransConfiguration configuration)
        {
            //var values = data as DataTable;
            var values = data.Rows;
            var rowIndex = Configuration.StartRowIndex;
            var maxColumnIndex = 0;
            foreach (DataRow item in values)
            {
                if (item is null) continue;
                var row = new Row() { RowIndex = ++rowIndex };
                var cellIndex = 0;
                foreach (var col in sheet.Columns)
                {
                    if (col.Ignore) continue;
                    var value = item[col.Name];
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