using LightExcel.OpenXml;
using LightExcel.Utils;

namespace LightExcel.Renders
{
    internal abstract class RenderBase : IDataRender
    {
        public abstract IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data, ExcelConfiguration configuration);

        public abstract IEnumerable<Row> RenderBody(object data, Sheet sheet, IEnumerable<ExcelColumnInfo> columns, ExcelConfiguration configuration);

        public virtual Row RenderHeader(IEnumerable<ExcelColumnInfo> columns, ExcelConfiguration configuration)
        {
            var row = new Row() { RowIndex = 1 };
            configuration.StartRowIndex = 1;
            var index = 0;
            foreach (var col in columns)
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