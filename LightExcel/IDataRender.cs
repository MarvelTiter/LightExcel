using DocumentFormat.OpenXml.Spreadsheet;
using LightExcel.OpenXml;

namespace LightExcel
{
    internal interface IDataRender
    {
        OpenXml.Row RenderHeader(object data);
        IEnumerable<OpenXml.Row> RenderBody(object data);
    }
}
