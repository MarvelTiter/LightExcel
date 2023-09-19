using LightExcel.OpenXml;

namespace LightExcel
{
    internal interface IDataRender
    {
        IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data, ExcelConfiguration configuration);
        Row RenderHeader(IEnumerable<ExcelColumnInfo> columns, ExcelConfiguration configuration);
        IEnumerable<Row> RenderBody(object data, Sheet sheet, IEnumerable<ExcelColumnInfo> columns, ExcelConfiguration configuration);
    }
}
