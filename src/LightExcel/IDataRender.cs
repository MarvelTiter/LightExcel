using LightExcel.OpenXml;
using LightExcel.OpenXml.Interfaces;

namespace LightExcel
{
    public interface IRenderSheet
    {
        ExcelColumnInfo[] Columns { get; set; }
        int MaxColumnIndex { get; set; }
        int MaxRowIndex { get; set; }
        void Write<TNode>(IEnumerable<TNode> children) where TNode : INode;
    }

    public interface IDataRender
    {
        IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data);
        internal void SetCustomHeaders(IEnumerable<Row> headers);
        internal IEnumerable<Row> RenderHeader(ExcelColumnInfo[] columns, TransConfiguration configuration);
        void Render(object data, IRenderSheet sheet, TransConfiguration configuration);
        internal IEnumerable<Row> RenderBody(object data, IRenderSheet sheet, TransConfiguration configuration);
        //void Render(object data, Sheet sheet, TransConfiguration configuration);
    }

#if NET6_0_OR_GREATER
    internal interface IAsyncDataRender
    {
        IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data);
        internal void SetCustomHeaders(IEnumerable<Row> headers);
        IEnumerable<Row> RenderHeader(ExcelColumnInfo[] columns, TransConfiguration configuration);
        IAsyncEnumerable<Row> RenderBodyAsync(object datas, Sheet sheet, TransConfiguration configuration, CancellationToken cancellationToken);
        Task RenderAsync(object datas, Sheet sheet, TransConfiguration configuration, CancellationToken cancellationToken);
    }
#endif
}
