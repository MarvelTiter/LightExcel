using LightExcel.Attributes;
using LightExcel.OpenXml;
using LightExcel.Utils;
using System.Collections;
using System.ComponentModel.DataAnnotations;
using System.Reflection;

namespace LightExcel.Renders
{
    internal class EnumerableEntityRender : RenderBase//, IDataRender
    {
        private readonly Type elementType;

        public EnumerableEntityRender(Type elementType, ExcelConfiguration configuration) : base(configuration)
        {
            this.elementType = elementType;
        }

        public override IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data)
        {
            var properties = elementType.GetProperties();
            int index = 1;
            foreach (var prop in properties)
            {
                ExcelColumnAttribute? excelColumnAttribute = prop.GetCustomAttribute<ExcelColumnAttribute>();
                DisplayAttribute? displayAttribute = prop.GetCustomAttribute<DisplayAttribute>();
                if (excelColumnAttribute?.Ignore ?? false) continue;
                var col = new ExcelColumnInfo(excelColumnAttribute?.Name ?? displayAttribute?.Name ?? prop.Name);
                col.Ignore = excelColumnAttribute?.Ignore ?? false;
                col.Property = new Property(prop);
                col.Type = prop.PropertyType;
                col.NumberFormat = excelColumnAttribute?.NumberFormat ?? false;
                col.Format = excelColumnAttribute?.Format;
                col.ColumnIndex = index++;
                col.AutoWidth = excelColumnAttribute?.AutoWidth ?? false;
                yield return col;
            }
        }

        public override IEnumerable<Row> RenderBody(object data, Sheet sheet, IEnumerable<ExcelColumnInfo> columns, TransConfiguration configuration)
        {
            var values = data as IEnumerable ?? throw new ArgumentException();
            var rowIndex = Configuration.StartRowIndex;
            var maxColumnIndex = 0;
            foreach (var item in values!)
            {
                if (item is null) continue;
                var row = new Row() { RowIndex = ++rowIndex };
                var cellIndex = 0;
                foreach (var col in columns)
                {
                    if (col.Property == null)
                    {
                        var p = elementType.GetProperty(col.Name);
                        if (p == null) continue;
                        col.Property = new Property(p);
                    }
                    var value = col.Property.GetValue(item);
                    cellIndex = col.ColumnIndex;
                    //var cell = new Cell();
                    //cell.Reference = ReferenceHelper.ConvertXyToCellReference(cellIndex, rowIndex);
                    //var (v, t) = CellHelper.FormatCell(value, Configuration, col);
                    ////cell.Type = CellHelper.ConvertCellType(col.Type);
                    ////cell.Value = CellHelper.GetCellValue(col, value, Configuration);
                    //cell.Value = v;
                    //cell.Type = t;
                    //cell.StyleIndex = col.NumberFormat || configuration.NumberFormatColumnFilter(col) ? "1" : null;
                    var nf = configuration.NumberFormatColumnFilter(col);
                    var cell = CellHelper.CreateCell(cellIndex, rowIndex, value, col, nf, Configuration);
                    row.AppendChild(cell);
                }
                maxColumnIndex = Math.Max(maxColumnIndex, cellIndex);
                yield return row;
            }
            sheet.MaxColumnIndex = maxColumnIndex;
            sheet.MaxRowIndex = rowIndex;
        }
    }
}