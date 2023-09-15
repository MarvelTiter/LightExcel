using LightExcel.OpenXml;
using System.Data;

namespace LightExcel.Renders
{
    internal class DataTableRender //: IDataRender
    {
        public IEnumerable<Row> RenderBody(object data)
        {
            var table = (DataTable)data;
            foreach (DataRow item in table.Rows)
            {
                var row = new Row();
                //foreach (DataColumn column in table.Columns)
                //{
                //    var cell = InternalHelper.CreateTypedCell(column.DataType, item[column]);
                //    row.AppendChild(cell);
                //}
                yield return row;
            }
        }

        public Row RenderHeader(object data)
        {
            var table = (DataTable)data;
            var row = new Row();
            foreach (DataColumn col in table.Columns)
            {
                //var cell = new Cell
                //{
                //    CellValue = new CellValue(col.ColumnName),
                //    DataType = new EnumValue<CellValues>(CellValues.String),
                //};
                //row.AppendChild(cell);
            }
            return row;
        }
    }
}