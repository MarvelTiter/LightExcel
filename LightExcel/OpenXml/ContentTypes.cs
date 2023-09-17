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


        protected override IEnumerable<INode> GetChildrenImpl(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.LocalName == "Default")
                {
                    var ext = reader["Extension"];
                    var ct = reader["ContentType"];
                    var def = new Default(ext, ct);
                    yield return def;
                }
                else if (reader.LocalName == "Override")
                {
                    var pn = reader["PartName"];
                    var ct = reader["ContentType"];
                    var ov = new Override(pn, ct);
                    yield return ov;
                }
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
