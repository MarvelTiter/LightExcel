using LightExcel.OpenXml;

namespace LightExcel
{
    internal interface IDataRender
    {
        void CollectExcelColumnInfo(object data, ExcelConfiguration configuration);
        Row RenderHeader(ExcelConfiguration configuration);
        IEnumerable<Row> RenderBody(object data, Sheet sheet, ExcelConfiguration configuration);
    }
}
