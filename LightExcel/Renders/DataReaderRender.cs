using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;

namespace LightExcel.Renders
{
    internal class DataReaderRender : IDataRender
    {
        private readonly WorkbookPart workbookPart;
        private readonly ExcelConfiguration configuration;

        public DataReaderRender(WorkbookPart workbookPart, ExcelConfiguration configuration)
        {
            this.workbookPart = workbookPart;
            this.configuration = configuration;
        }

        public IEnumerable<Row> RenderBody(object data)
        {
            var reader = (IDataReader)data;
            //int rowValueIndex = 0;
            while (reader.Read())
            {
                var row = new Row();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var field = reader.GetName(i);
                    var value = reader.GetValue(i);
                    var cell = InternalHelper.CreateTypedCell(reader.GetFieldType(i), value);
                    if (configuration.HasStyle(field, value))
                    {
                        cell.StyleIndex = configuration.GetStyleIndex(field, workbookPart);
                    }
                    row.AppendChild(cell);
                }
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
                var cell = new Cell
                {
                    CellValue = new CellValue(reader.GetName(i)),
                    DataType = new EnumValue<CellValues>(CellValues.String),
                    //CellReference = $"Header{i}"
                };
                row.AppendChild(cell);
            }
            return row;
        }
    }
}