using LightExcel.OpenXml;
using System.Data;

namespace LightExcel.Renders
{
    internal class DataReaderRender : IDataRender
    {
        public IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data, ExcelConfiguration configuration)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<Row> RenderBody(object data, Sheet sheet, IEnumerable<ExcelColumnInfo> columns, ExcelConfiguration configuration)
        {
            throw new NotImplementedException();
        }

        public Row RenderHeader(IEnumerable<ExcelColumnInfo> columns, ExcelConfiguration configuration)
        {
            throw new NotImplementedException();
        }
    }
}