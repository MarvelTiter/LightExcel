using LightExcel.OpenXml;
using System.Data;

namespace LightExcel.Renders
{
    internal class DataReaderRender : IDataRender
    {
        public void CollectExcelColumnInfo(object data, ExcelHelperConfiguration configuration)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<Row> RenderBody(object data, Sheet sheet, ExcelHelperConfiguration configuration)
        {
            throw new NotImplementedException();
        }

        public Row RenderHeader(ExcelHelperConfiguration configuration)
        {
            throw new NotImplementedException();
        }
    }
}