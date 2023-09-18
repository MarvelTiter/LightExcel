using LightExcel.OpenXml;
using System.Data;

namespace LightExcel.Renders
{
    internal class DataReaderRender : IDataRender
    {
        public void CollectExcelColumnInfo(object data, ExcelConfiguration configuration)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<Row> RenderBody(object data, Sheet sheet, ExcelConfiguration configuration)
        {
            throw new NotImplementedException();
        }

        public Row RenderHeader(ExcelConfiguration configuration)
        {
            throw new NotImplementedException();
        }
    }
}