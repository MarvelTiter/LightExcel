using System.IO.Compression;
using System.Text;
using System.Xml;

namespace LightExcel.OpenXml
{
    internal abstract class Node
    {
        internal abstract string ToXmlString();
    }
    internal class Override : Node
    {
        public Override(string? PartName, string? ContentType)
        {
            this.PartName = PartName;
            this.ContentType = ContentType;
        }

        public string? PartName { get; }
        public string? ContentType { get; }
        internal override string ToXmlString()
        {
            return ($"<Override PartName=\"{PartName}\" /> ContentType=\"{ContentType}\" />");
        }
    }
    internal class Default : Node
    {
        public Default(string? extension, string? contentType)
        {
            Extension = extension;
            ContentType = contentType;
        }

        public string? Extension { get; }
        public string? ContentType { get; }

        internal override string ToXmlString()
        {
            return $"<Default Extension=\"{Extension}\" ContentType=\"{ContentType}\" />";
        }
    }
    /// <summary>
    /// [Content_Types].xml
    /// </summary>
    internal class ContentTypes : XmlPart<Node>
    {
        public ContentTypes(ZipArchive archive) : base(archive)
        {

        }

        internal void AppendChild(Node child)
        {
            Children!.Add(child);
        }

        protected override IEnumerable<Node> GetChildren()
        {
            if (reader == null)
            {
                yield break;
            }
            while (reader.Read())
            {
                if (reader.LocalName == "Default")
                {
                    var ext = reader["Extension"];
                    var ct = reader["ContentType"];
                    var def = new Default(ext, ct);
                    Children.Add(def);
                    yield return def;
                }
                else if (reader.LocalName == "Override")
                {
                    var pn = reader["PartName"];
                    var ct = reader["ContentType"];
                    var ov = new Override(pn, ct);
                    Children.Add(ov);
                    yield return ov;
                }
            }
        }

        internal void ToXmlString(StringBuilder builder)
        {
            //builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            //builder.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            //builder.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />");
            //builder.Append("<Default Extension=\"xml\" ContentType=\"application/xml\" />");
            //foreach (Node node in children)
            //{
            //    node.ToXmlString();
            //}
            //builder.Append("</Types>");
        }

        internal override void Save()
        {
            throw new NotImplementedException();
        }
    }
}
