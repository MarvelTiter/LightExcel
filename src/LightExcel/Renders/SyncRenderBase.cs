using LightExcel.OpenXml;
using System.Reflection;

namespace LightExcel.Renders
{
    internal abstract class SyncRenderBase<T, TElement>(ExcelConfiguration configuration)
        : RenderBase(configuration), IDataRender
    {
        public IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data)
            => CollectExcelColumnInfo(GetFirstElement((T)data));


        public IEnumerable<Row> RenderBody(object data, IRenderSheet sheet, TransConfiguration configuration)
            => RenderBody((T)data, sheet, configuration);

        public abstract IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(TElement data);
        public abstract IEnumerable<Row> RenderBody(T data, IRenderSheet sheet, TransConfiguration configuration);
        public abstract TElement GetFirstElement(T data);

        public void Render(object data, IRenderSheet sheet, TransConfiguration configuration)
        {
            if (sheet.Columns.Length == 0)
            {
                ExcelColumnInfo[] columns = [.. CollectExcelColumnInfo(data)];
                sheet.Columns = columns;
            }
            var allRows = CollectALlRows(data, sheet, sheet.Columns, configuration);
            sheet.Write(allRows);
        }

        private IEnumerable<Row> CollectALlRows(object data, IRenderSheet sheet, ExcelColumnInfo[] columns, TransConfiguration configuration)
        {
            if (Configuration.UseHeader)
            {
                var headers = RenderHeader(columns, configuration);
                foreach (var row in headers)
                {
                    yield return row;
                }
            }
            var datas = RenderBody(data, sheet, configuration);
            foreach (var row in datas)
            {
                yield return row;
            }
        }
    }
}