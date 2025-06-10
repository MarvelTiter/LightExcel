using System.Text;
using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;

namespace LightExcel.OpenXml
{
    internal readonly struct Cell : INode
    {
        //public Cell(){}
        public Cell(string? value, string? reference, string? type)
            : this(value, reference, type, null, null)
        {

        }
        public Cell(string? value, string? reference, string? type, string? style)
            : this(value, reference, type, style, null)
        {

        }
        public Cell(string? value, string? reference, string? type, string? style, int? columnIndex)
        {
            Value = value;
            Reference = reference;
            Type = type;
            StyleIndex = style;
            ColumnIndex = columnIndex;
        }
        // public Cell(string? value, string? reference, string? type, string? style, int? index): this(value, reference, type, style)
        // {
        //     ColumnIndex = index;
        // }
        public string? Reference { get; }
        public string? Type { get; }
        public string? StyleIndex { get; }
        public string? Value { get; }
        public int? ColumnIndex { get; }

        public void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write($"<c r=\"{Reference}\" {(Type == null ? "" : $"t=\"{Type}\"")} {(StyleIndex != null ? $"s=\"{StyleIndex}\"" : "")}>");
            writer.Write($"<v>{Value}</v>");
            writer.Write("</c>");
        }

        public static Cell EmptyCell(int x, int y)
        {
            var r = ReferenceHelper.ConvertXyToCellReference(x, y);
            return new Cell(null, r, null, null, null);
        }
    }
}
