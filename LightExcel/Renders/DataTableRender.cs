using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;

namespace LightExcel.Renders
{
    internal class DataTableRender : IDataRender
    {
        private readonly WorkbookPart workbookPart;
        private readonly ExcelConfiguration configuration;

        public DataTableRender(WorkbookPart workbookPart, ExcelConfiguration configuration)
        {
            this.workbookPart = workbookPart;
            this.configuration = configuration;
        }

        public IEnumerable<Row> RenderBody(object data)
        {
            var table = (DataTable)data;
            foreach (DataRow item in table.Rows)
            {
                var row = new Row();
                foreach (DataColumn column in table.Columns)
                {
                    object value = item[column];
                    var cell = InternalHelper.CreateTypedCell(column.DataType, value);
                    if (configuration.HasStyle(column.ColumnName, value))
                    {
                        cell.StyleIndex = configuration.GetStyleIndex(column.ColumnName, workbookPart);
                    }
                    row.AppendChild(cell);
                }
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
                };
                row.AppendChild(cell);
            }
            return row;
        }
    }
}