using DocumentFormat.OpenXml.Spreadsheet;

namespace LightExcel
{
    public interface IDataRender
    {
        Row RenderHeader(object data);
        IEnumerable<Row> RenderBody(object data);
    }
}
