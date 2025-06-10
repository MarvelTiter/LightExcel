using LightExcel.OpenXml;
using LightExcel.Utils;

namespace LightExcel.Renders;
internal abstract class RenderBase(ExcelConfiguration configuration)
{
    public ExcelConfiguration Configuration { get; } = configuration;
    IEnumerable<Row>? customHeaders;
    public virtual void SetCustomHeaders(IEnumerable<Row> headers)
    {
        customHeaders = headers;
    }

    public virtual IEnumerable<Row> RenderHeader(ExcelColumnInfo[] columns, TransConfiguration configuration)
    {
        if (customHeaders is not null)
        {
            foreach (var item in customHeaders)
            {
                yield return item;
            }
            yield break;
        }
        var row = new Row() { RowIndex = 1 };
        Configuration.StartRowIndex = 1;
        var index = 0;
        foreach (var col in columns)
        {
            var r = ReferenceHelper.ConvertXyToCellReference(++index, 1);
            var cell = new Cell(col.Name, r, "str");
            //{
            //    Reference = ReferenceHelper.ConvertXyToCellReference(++index, 1),
            //    Type = "str",
            //    Value = col.Name
            //};
            row.AppendChild(cell);
        }
        yield return row;
    }

    protected void AssignDynamicInfo(ExcelColumnInfo origin)
    {
        var dyCol = Configuration[origin.Name];
        origin.Format ??= dyCol?.Format;
        origin.Width ??= dyCol?.Width;
        origin.AutoWidth = !origin.Width.HasValue && (dyCol?.AutoWidth ?? Configuration.AutoWidth);
    }
}