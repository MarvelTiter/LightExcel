using LightExcel.Attributes;
using LightExcel.OpenXml;
using LightExcel.Utils;
using System.Collections;
using System.Reflection;

namespace LightExcel.Renders
{
    internal class EnumerableEntityRender : RenderBase, IDataRender
    {
        private readonly Type elementType;
        private bool renderHeader;

        public EnumerableEntityRender(Type elementType)
        {
            this.elementType = elementType;
        }

        public void CollectExcelColumnInfo(object data, ExcelHelperConfiguration configuration)
        {
            var properties = elementType.GetProperties();
            foreach (var prop in properties)
            {
                ExcelColumnAttribute? excelColumnAttribute = prop.GetCustomAttribute<ExcelColumnAttribute>();
                var col = new ExcelColumnInfo(excelColumnAttribute?.Name ?? prop.Name);
                col.Ignore = excelColumnAttribute?.Ignore ?? false;
                col.Property = new Property(prop);
                Columns.Add(col);
            }
        }

        public IEnumerable<Row> RenderBody(object data, Sheet sheet, ExcelHelperConfiguration configuration)
        {
            var values = data as IEnumerable;
            var rowIndex = configuration.UseHeader ? 1 : 0;
            foreach (var item in values!)
            {
                if (item is null) continue;
                var row = new Row() { RowIndex = ++rowIndex };
                var cellIndex = 0;
                foreach (var col in Columns)
                {
                    if (col.Ignore) continue;
                    var cell = new Cell();
                    var value = col.Property!.GetValue(item);
                    cell.Reference = ReferenceHelper.ConvertXyToCellReference(++cellIndex, rowIndex);
                    cell.Type = CellHelper.ConvertCellType(col.Property!.Info.PropertyType);
                    cell.Value = CellHelper.GetCellValue(col, value, configuration);
                    row.AppendChild(cell);
                }
                yield return row;
            }

        }

        public Row RenderHeader(ExcelHelperConfiguration configuration)
        {
            var row = new Row() { RowIndex = 1 };
            var index = 0;
            foreach (var col in Columns)
            {
                var cell = new Cell
                {
                    Reference = ReferenceHelper.ConvertXyToCellReference(++index, 1),
                    Type = "str",
                    Value = col.Name
                };
                row.AppendChild(cell);
            }
            return row;
        }
    }
}