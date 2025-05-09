using System.Collections;
using LightExcel.OpenXml.Basic;
using LightExcel.OpenXml.Interfaces;

namespace LightExcel.OpenXml.Styles
{
    internal class BorderCollection : SimpleNodeCollectionXmlPart<Border>
    {
        public override void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write($"<borders count={Count}>");
            foreach (var item in Children)
            {
                item.WriteToXml(writer);
            }
            writer.Write("</borders>");
        }
    }

    internal class Border : IStyleNode
    {
        public int Id { get; }
        public void WriteToXml(LightExcelStreamWriter writer)
        {

        }
    }
}
