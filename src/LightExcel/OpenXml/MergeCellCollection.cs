using LightExcel.OpenXml.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.OpenXml
{
    internal class MergeCellCollection : INodeCollection<MergeCell>, INode
    {
        public int Count => _cells.Count;
        IList<MergeCell> _cells = new List<MergeCell>();
        public void AppendChild(MergeCell child)
        {
            _cells.Add(child);
        }

        public void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write($"<mergeCells count=\"{Count}\">");
            foreach (var item in _cells)
            {
                item.WriteToXml(writer);
            }
            writer.Write("</mergeCells>");
        }
    }

    internal class MergeCell : INode
    {
        public string MergeRef { get; set; }
        public MergeCell(string mergeref)
        {
            MergeRef = mergeref;
        }
        public void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write($"<mergeCell ref=\"{MergeRef}\" />");
        }
    }
}
