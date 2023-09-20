using LightExcel.OpenXml.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.OpenXml
{
    internal class Col : INode
    {
        public int Min { get; set; }
        public int Max { get; set; }
        public double? Width { get; set; }
        public bool CustomWidth { get; set; }
        string CwString => CustomWidth ? "customWidth=\"1\"" : "";
        string WString => Width.HasValue ? $"width=\"{Width}\"" : "";
        public void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write($"<col min=\"{Min}\" max=\"{Max}\" {WString} {CwString} />");
        }
    }
    internal class SheetCols : INodeCollection<Col>,INode
    {
        public int Count => Cols.Count;
        List<Col> Cols = new List<Col>();
        public void AppendChild(Col child)
        {
            Cols.Add(child);
        }

        public void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write("<cols xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            foreach (var item in Cols)
            {
                item.WriteToXml(writer);
            }
            writer.Write("</cols>");
        }
    }
}
