using System.Text;
using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;

namespace LightExcel.OpenXml
{
    internal struct Cell : INode
    {
        public Cell(){}
        // public Cell(string? value, string? reference, string? type, string? style)
        // {
        //     Value = value;
        //     Reference = reference;
        //     Type = type;
        //     StyleIndex = style;
        // }
        // public Cell(string? value, string? reference, string? type, string? style, int? index): this(value, reference, type, style)
        // {
        //     ColumnIndex = index;
        // }
        public string? Reference { get; set; }
        public string? Type { get; set; }
        public string? StyleIndex { get; set; }
        public string? Value { get; set; }
        public int? ColumnIndex { get; set; }

        public void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write($"<c r=\"{Reference}\" {(Type == null ? "": $"t=\"{Type}\"")} {(StyleIndex != null ? $"s=\"{StyleIndex}\"":"")}>");
            writer.Write($"<v>{Value}</v>");
            writer.Write("</c>");
        }

        public static Cell EmptyCell(int x, int y)
        {
            var r = ReferenceHelper.ConvertXyToCellReference(x, y);
            return new Cell
            {
                Reference = r
            };
        }
    }
}
