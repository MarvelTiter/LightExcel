using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;
using System.Collections;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace LightExcel.OpenXml
{
    internal class RelationshipCollection : NodeCollectionXmlPart<Relationship>
    {
        public RelationshipCollection(ZipArchive archive) : base(archive, "xl/_rels/workbook.xml.rels")
        {

        }

        protected override IEnumerable<Relationship> GetChildrenImpl(LightExcelXmlReader reader)
        {
            if (!reader.IsStartWith("Relationships", XmlHelper.RelaNs)) yield break;
            if (!reader.ReadFirstContent()) yield break;
            while (!reader.EOF)
            {
                if (reader.IsStartWith("Relationship", XmlHelper.RelaNs))
                {
                    var id = reader.GetAttribute("Id") ?? throw new Exception("Excel Xml Relationship Error (without id)");
                    var type = reader.GetAttribute("Type") ?? throw new Exception("Excel Xml Relationship Error (without type)");
                    var target = reader.GetAttribute("Target") ?? throw new Exception("Excel Xml Relationship Error (without target)");
                    var rel = new Relationship(id, type, target);
                    yield return rel;
                }
            }
        }

        protected override void WriteImpl(LightExcelStreamWriter writer, IEnumerable<INode> children)
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            writer.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            foreach (var child in children)
            {
                child.WriteToXml(writer);
            }
            writer.Write("</Relationships>");
        }

    }
}
