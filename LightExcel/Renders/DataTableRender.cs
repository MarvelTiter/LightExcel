using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;

namespace LightExcel.Renders
{
    internal class DataTableRender : IDataRender
    {
        public IEnumerable<Row> RenderBody(object data)
        {
            var table = (DataTable)data;
            int rowValueIndex = 0;
            foreach (DataRow item in table.Rows)
            {
                var row = new Row();
                foreach (DataColumn column in table.Columns)
                {
                    var cell = InternalHelper.CreateTypedCell(column.DataType, item[column]);
                    cell.CellReference = $"{column.ColumnName}{rowValueIndex}";
                    row.AppendChild(cell);
                }
                rowValueIndex++;
                yield return row;
            }
        }

        public Row RenderHeader(object data)
        {
            var table = (DataTable)data;
            var row = new Row();
            foreach (DataColumn col in table.Columns)
            {
                var cell = new Cell
                {
                    CellValue = new CellValue(col.ColumnName),
                    DataType = new EnumValue<CellValues>(CellValues.String),
                    CellReference = $"Header{col.ColumnName}"
                };
                row.AppendChild(cell);
            }
            return row;
        }
    }
}