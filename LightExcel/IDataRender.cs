using LightExcel.OpenXml;

namespace LightExcel
{
    internal interface IDataRender
    {
        Row RenderHeader(Sheet sheet, ExcelHelperConfiguration configuration);
        IEnumerable<Row> RenderBody(object data, Sheet sheet, ExcelHelperConfiguration configuration);
    }
}
