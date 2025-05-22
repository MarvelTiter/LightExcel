using LightExcel.OpenXml.Interfaces;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LightExcel.OpenXml.Basic;

namespace LightExcel.OpenXml
{
    internal class MergeCellCollection : SimpleNodeCollectionXmlPart<MergeCell>
    {
        public override void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write($"<mergeCells count=\"{Count}\">");
            foreach (var item in Children)
            {
                item.WriteToXml(writer);
            }
            writer.Write("</mergeCells>");
        }
    }

    internal class MergeCell(string mergeref) : INode
    {
        public string MergeRef { get; set; } = mergeref;

        public void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write($"<mergeCell ref=\"{MergeRef}\" />");
        }
    }
}
