using LightExcel.OpenXml;

namespace LightExcel
{
    internal interface IDataRender
    {
        IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data);
        Row RenderHeader(ExcelColumnInfo[] columns, TransConfiguration configuration);
        IEnumerable<Row> RenderBody(object data, Sheet sheet, ExcelColumnInfo[] columns, TransConfiguration configuration);
    }
}
