using LightExcel.OpenXml;

namespace LightExcel
{
    internal interface IDataRender
    {
        void CollectExcelColumnInfo(object data, ExcelHelperConfiguration configuration);
        Row RenderHeader(ExcelHelperConfiguration configuration);
        IEnumerable<Row> RenderBody(object data, Sheet sheet, ExcelHelperConfiguration configuration);
    }
}
