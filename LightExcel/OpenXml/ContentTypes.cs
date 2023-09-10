using System.Text;

namespace LightExcel.OpenXml
{
    internal abstract class Node
    {
        internal virtual void AppendChild(Node node)
        {

        }

        internal abstract void ToXmlString(StringBuilder builder);
    }
    internal class Override : Node
    {
        public Override(string PartName, string ContentType)
        {
            this.PartName = PartName;
            this.ContentType = ContentType;
        }

        public string PartName { get; }
        public string ContentType { get; }

        internal override void ToXmlString(StringBuilder builder)
        {
            builder.Append($"<Override PartName=\"/{PartName}\" /> ContentType=\"{ContentType}\"");
        }
    }
    internal class ContentTypes : Node
    {
        readonly IList<Node> children = new List<Node>();
        internal override void AppendChild(Node node)
        {
            children.Add(node);
        }

        internal override void ToXmlString(StringBuilder builder)
        {
            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            foreach (Node node in children)
            {
                node.ToXmlString(builder);
            }
            builder.Append("</Types>");
        }
    }
}
