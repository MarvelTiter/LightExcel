using System.Text;
using LightExcel.OpenXml.Interfaces;

namespace LightExcel.OpenXml
{
    internal class Cell : INode
    {
        public string? Reference { get; set; }
        public string? Type { get; set; }
        public string? StyleIndex { get; set; }
        public string? Value { get; set; }

        public void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write($"<c r=\"{Reference}\" t=\"{Type}\" {(StyleIndex != null ? $"s=\"{StyleIndex}\"":"")}>");
            writer.Write($"<v>{Value}</v>");
            writer.Write("</c>");
        }
    }
}
