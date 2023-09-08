using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using LightExcel.Attributes;
using System.Collections;
using System.Reflection;

namespace LightExcel.Renders
{
    internal class EnumerableEntityRender : IDataRender
    {
        private readonly Type elementType;
        private readonly WorkbookPart workbookPart;
        private readonly ExcelConfiguration configuration;
        private readonly PropertyInfo[] properties;
        private readonly Dictionary<string, PropertyInfo> validProp;
        public EnumerableEntityRender(Type elementType, WorkbookPart workbookPart, ExcelConfiguration configuration)
        {
            this.elementType = elementType;
            this.workbookPart = workbookPart;
            this.configuration = configuration;
            properties = elementType.GetProperties();
            validProp = new Dictionary<string, PropertyInfo>();
            foreach (var prop in properties)
            {
                ExcelColumnAttribute? excelColumnAttribute = prop.GetCustomAttribute<ExcelColumnAttribute>();
                if (excelColumnAttribute?.Ignore ?? false) continue;
                validProp.Add(excelColumnAttribute?.Name ?? prop.Name, prop);
            }
        }
        public IEnumerable<Row> RenderBody(object data)
        {
            var values = data as IEnumerable;
            foreach (var item in values!)
            {
                if (item is null) continue;
                var row = new Row();
                foreach (var kv in validProp)
                {
                    var prop = kv.Value;
                    var value = prop!.GetValue(item);
                    var cell = InternalHelper.CreateTypedCell(prop.PropertyType, prop!.GetValue(item) ?? "");
                    if (configuration.HasStyle(kv.Key, value))
                    {
                        cell.StyleIndex = configuration.GetStyleIndex(kv.Key, workbookPart);
                    }
                    row.AppendChild(cell);
                }
                yield return row;
            }

        }

        public Row RenderHeader(object data)
        {
            var row = new Row();
            foreach (var kv in validProp)
            {
                var cell = new Cell
                {
                    CellValue = new CellValue(kv.Key),
                    DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String),
                };
                row.AppendChild(cell);
            }
            return row;
        }
    }
}