using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;

namespace LightExcel.Renders
{
    internal class DataReaderRender : IDataRender
    {
        public IEnumerable<Row> RenderBody(object data)
        {
            var reader = (IDataReader)data;
            while (reader.Read())
            {
                var row = new Row();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var cell = InternalHelper.CreateTypedCell(reader.GetFieldType(i), reader.GetValue(i));
                    row.AppendChild(cell);
                }
                yield return row;
            }
        }

        public Row RenderHeader(object data)
        {
            var reader = (IDataReader)data;
            var row = new Row();
            for (int i = 0; i < reader.FieldCount; i++)
            {
                var cell = new Cell
                {
                    CellValue = new CellValue(reader.GetName(i)),
                    DataType = new EnumValue<CellValues>(CellValues.String),
                };
                row.AppendChild(cell);
            }
            return row;
        }
    }
}