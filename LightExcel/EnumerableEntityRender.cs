using DocumentFormat.OpenXml.Spreadsheet;
using LightExcel.Attributes;
using System.Collections;
using System.Reflection;

namespace LightExcel
{
    internal class EnumerableEntityRender : IDataRender
    {
        private readonly Type elementType;
        private readonly PropertyInfo[] properties;
        public EnumerableEntityRender(Type elementType)
        {
            this.elementType = elementType;
            this.properties = elementType.GetProperties();
        }
        public IEnumerable<Row> RenderBody(object data)
        {
            var values = data as IEnumerable;
            foreach (var item in values!)
            {
                var row = new Row();
                foreach (var prop in properties)
                {
                    var cell = InternalHelper.CreateTypedCell(prop.PropertyType, prop.GetValue(item));
                    row.AppendChild(cell);
                }
                yield return row;
            }

        }

        public Row RenderHeader(object data)
        {
            var row = new Row();
            foreach (var prop in properties)
            {
                var colName = prop.GetCustomAttribute<ExcelColumnAttribute>()?.Name ?? prop.Name;
                var cell = new Cell
                {
                    CellValue = new CellValue(colName),
                    DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String),
                };
                row.AppendChild(cell);
            }
            return row;
        }
    }
}