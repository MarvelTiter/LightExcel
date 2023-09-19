using DocumentFormat.OpenXml.Spreadsheet;
using LightExcel.OpenXml.Interfaces;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace LightExcel.OpenXml
{
    internal class Row : INodeCollection<Cell>, INode
    {
        public int RowIndex { get; set; }
        public bool IsTemplateRow { get; set; }
        public List<Cell> RowDatas { get; set; } = new List<Cell>();

        public int Count => RowDatas.Count;

        public void AppendChild(Cell child)
        {
            RowDatas.Add(child);
        }

        public void WriteToXml(LightExcelStreamWriter writer)
        {
            if (IsTemplateRow) { return; }
            writer.Write($"<row r=\"{RowIndex}\">");
            foreach (Cell cell in RowDatas)
            {
                cell.WriteToXml(writer);
            }
            writer.Write("</row>");
        }
    }
}
