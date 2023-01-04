using DocumentFormat.OpenXml.Spreadsheet;

namespace LightExcel
{
    internal class EnumerableEntityRender : IDataRender
    {
        public IEnumerable<Row> RenderBody(object data)
        {
            throw new NotImplementedException();
        }

        public Row RenderHeader(object data)
        {
            throw new NotImplementedException();
        }
    }
}