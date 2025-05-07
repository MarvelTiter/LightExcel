using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace LightExcel.OpenXml
{
    /// <summary>
    /// [Content_Types].xml
    /// </summary>
    internal class ContentTypes : NodeCollectionXmlPart<INode>
    {
        public ContentTypes(ZipArchive archive) : base(archive, "[Content_Types].xml")
        {

        }


        protected override IEnumerable<INode> GetChildrenImpl(LightExcelXmlReader reader)
        {
            if (!reader.IsStartWith("Types", XmlHelper.ContentTypesXmlns))
                yield break;
            if (!reader.ReadFirstContent()) yield break;
            while (!reader.EOF)
            {
                if (reader.IsStartWith("Override", XmlHelper.ContentTypesXmlns))
                {
                    var pn = reader.GetAttribute("PartName");
                    var ct = reader.GetAttribute("ContentType");
                    var ov = new Override(pn, ct);
                    yield return ov;
                }
                else if (!reader.SkipContent()) 
                    yield break;
            }
        }

        protected override void WriteImpl(LightExcelStreamWriter writer, IEnumerable<INode> children)
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            writer.Write("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            writer.Write("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />");
            writer.Write("<Default Extension=\"xml\" ContentType=\"application/xml\" />");
            foreach (var node in children)
            {
                node.WriteToXml(writer);
            }
            writer.Write("</Types>");
        }
    }
}
