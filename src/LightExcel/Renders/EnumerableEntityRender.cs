using LightExcel.Attributes;
using LightExcel.OpenXml;
using LightExcel.Utils;
using System.Collections;
using System.Reflection;

namespace LightExcel.Renders
{
    internal class EnumerableEntityRender<T> : SyncRenderBase<IEnumerable<T>, T>//, IDataRender
    {
        private readonly Type elementType;

        public EnumerableEntityRender(ExcelConfiguration configuration) : base(configuration)
        {
            this.elementType = typeof(T);
        }

        public override IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(T data)
        {
            //            var properties = elementType.GetProperties();
            //            int index = 1;
            //            foreach (var prop in properties)
            //            {
            //                ExcelColumnAttribute? excelColumnAttribute = prop.GetCustomAttribute<ExcelColumnAttribute>();
            //                if (excelColumnAttribute?.Ignore ?? false) continue;
            //#if NET6_0_OR_GREATER
            //                var displayAttribute = prop.GetCustomAttribute<System.ComponentModel.DataAnnotations.DisplayAttribute>();
            //                var col = new ExcelColumnInfo(excelColumnAttribute?.Name ?? displayAttribute?.Name ?? prop.Name);
            //#else
            //                var col = new ExcelColumnInfo(excelColumnAttribute?.Name ??  prop.Name);
            //#endif

            //                col.Ignore = excelColumnAttribute?.Ignore ?? false;
            //                col.Property = new Property(prop);
            //                col.Type = prop.PropertyType;
            //                col.NumberFormat = excelColumnAttribute?.NumberFormat ?? false;
            //                col.Format = excelColumnAttribute?.Format;
            //                col.ColumnIndex = index++;
            //                col.AutoWidth = excelColumnAttribute?.AutoWidth ?? false;
            //                col.Width = excelColumnAttribute?.Width;
            //                AssignDynamicInfo(col);
            //                yield return col;
            //            }
            return elementType.CollectEntityInfo(AssignDynamicInfo);
        }

        public override T GetFirstElement(IEnumerable<T> data) => data.First();

        public override IEnumerable<Row> RenderBody(IEnumerable<T> data, IRenderSheet sheet, TransConfiguration configuration)
        {
            //var values = data as IEnumerable ?? throw new ArgumentException();
            var values = data;
            var rowIndex = Configuration.StartRowIndex;
            var maxColumnIndex = 0;
            foreach (var item in values)
            {
                if (item is null) continue;
                var row = new Row() { RowIndex = ++rowIndex };
                var cellIndex = 0;
                foreach (var col in sheet.Columns)
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
                    var cell = CellHelper.CreateCell(cellIndex, rowIndex, value, col, configuration);
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