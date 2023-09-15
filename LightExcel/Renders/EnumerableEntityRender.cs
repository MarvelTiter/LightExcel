using LightExcel.Attributes;
using LightExcel.OpenXml;
using LightExcel.Utils;
using System.Collections;
using System.Reflection;

namespace LightExcel.Renders
{
    internal class EnumerableEntityRender : IDataRender
    {
        private readonly Type elementType;
        private readonly PropertyInfo[] properties;
        private readonly Dictionary<string, PropertyInfo> validProp;
        public EnumerableEntityRender(Type elementType)
        {
            this.elementType = elementType;
            properties = elementType.GetProperties();
            validProp = new Dictionary<string, PropertyInfo>();
            foreach (var prop in properties)
            {
                ExcelColumnAttribute? excelColumnAttribute = prop.GetCustomAttribute<ExcelColumnAttribute>();
                if (excelColumnAttribute?.Ignore ?? false) continue;
                validProp.Add(excelColumnAttribute?.Name ?? prop.Name, prop);
            }
        }
        public IEnumerable<Row> RenderBody(object data, Sheet sheet, ExcelHelperConfiguration configuration)
        {
            var values = data as IEnumerable;
            foreach (var item in values!)
            {
                if (item is null) continue;
                var row = new Row();
                //foreach (var kv in validProp)
                //{
                //    var prop = kv.Value;
                //    var cell = InternalHelper.CreateTypedCell(prop.PropertyType, prop!.GetValue(item) ?? "");
                //    row.AppendChild(cell);
                //}
                yield return row;
            }

        }

        public Row RenderHeader(Sheet sheet, ExcelHelperConfiguration configuration)
        {
            var row = new Row() { RowIndex = 1 };
            var index = 0;
            foreach (var kv in validProp)
            {
                row.RowDatas.Add(new Cell
                {
                    Reference = ReferenceHelper.ConvertXyToCellReference(++index, 1)
                });
            }
            return row;
        }
    }
}