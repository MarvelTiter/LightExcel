using LightExcel.OpenXml;
using System.Data;

namespace LightExcel.Renders
{
    internal class DataReaderRender //: IDataRender
    {
        public IEnumerable<Row> RenderBody(object data)
        {
            var reader = (IDataReader)data;
            //int rowValueIndex = 0;
            while (reader.Read())
            {
                var row = new Row();
                //for (int i = 0; i < reader.FieldCount; i++)
                //{
                //    var cell = InternalHelper.CreateTypedCell(reader.GetFieldType(i), reader.GetValue(i));
                //    //cell.CellReference = $"{reader.GetName(i)}{rowValueIndex}";
                //    row.AppendChild(cell);
                //}
                //rowValueIndex++;
                yield return row;
            }
        }

        public Row RenderHeader(object data)
        {
            var reader = (IDataReader)data;
            var row = new Row();
            for (int i = 0; i < reader.FieldCount; i++)
            {
                //var cell = new Cell
                //{
                //    CellValue = new CellValue(reader.GetName(i)),
                //    DataType = new EnumValue<CellValues>(CellValues.String),
                //    //CellReference = $"Header{i}"
                //};
                //row.AppendChild(cell);
            }
            return row;
        }
    }
}