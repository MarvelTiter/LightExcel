using LightExcel.Attributes;
using LightExcel.OpenXml;
using LightExcel.Utils;
using System.Collections;
using System.Reflection;

namespace LightExcel.Renders
{
    internal class EnumerableEntityRender : RenderBase//, IDataRender
    {
        private readonly Type elementType;

        public EnumerableEntityRender(Type elementType)
        {
            this.elementType = elementType;
        }

        public override IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data, ExcelConfiguration configuration)
        {
            var properties = elementType.GetProperties();
            foreach (var prop in properties)
            {
                ExcelColumnAttribute? excelColumnAttribute = prop.GetCustomAttribute<ExcelColumnAttribute>();
                var col = new ExcelColumnInfo(excelColumnAttribute?.Name ?? prop.Name);
                col.Ignore = excelColumnAttribute?.Ignore ?? false;
                col.Property = new Property(prop);
                col.NumberFormat = excelColumnAttribute?.NumberFormat ?? false;
                yield return col;
            }
        }

        public override IEnumerable<Row> RenderBody(object data, Sheet sheet, IEnumerable<ExcelColumnInfo> columns, ExcelConfiguration configuration)
        {
            var values = data as IEnumerable;
            var rowIndex = configuration.StartRowIndex;
            foreach (var item in values!)
            {
                if (item is null) continue;
                var row = new Row() { RowIndex = ++rowIndex };
                var cellIndex = 0;
                foreach (var col in columns)
                {
                    if (col.Ignore) continue;
                    if (col.Property == null)
                    {
                        var p = elementType.GetProperty(col.Name);
                        if (p == null) continue;
                        col.Property = new Property(p);
                    }
                    var cell = new Cell();
                    var value = col.Property.GetValue(item);
                    cell.Reference = ReferenceHelper.ConvertXyToCellReference(++cellIndex, rowIndex);
                    cell.Type = CellHelper.ConvertCellType(col.Property!.Info.PropertyType);
                    cell.Value = CellHelper.GetCellValue(col, value, configuration);
                    cell.StyleIndex = col.NumberFormat ? "1" : null;
                    row.AppendChild(cell);
                }
                yield return row;
            }

        }
    }
}