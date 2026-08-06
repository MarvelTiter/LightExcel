using LightExcel.OpenXml.Basic;
using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Text;
using System.Xml;

namespace LightExcel.OpenXml
{
    /// <summary>
    /// [Content_Types].xml
    /// </summary>
    internal class ContentTypes(ZipArchive archive, ExcelConfiguration configuration)
        : NodeCollectionXmlPart<INode>(archive, "[Content_Types].xml", configuration)
    {
        protected override IEnumerable<INode> GetChildrenImpl()
        {
            if (reader is null) yield break;
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
#if NET8_0_OR_GREATER
        protected override async IAsyncEnumerable<INode> GetChildrenImplAsync([EnumeratorCancellation] CancellationToken cancellationToken)
        {
            if (reader is null) yield break;
            if (!reader.IsStartWith("Types", XmlHelper.ContentTypesXmlns))
                yield break;
            if (!await reader.ReadFirstContentAsync()) yield break;
            while (!reader.EOF)
            {
                if (reader.IsStartWith("Override", XmlHelper.ContentTypesXmlns))
                {
                    var pn = reader.GetAttribute("PartName");
                    var ct = reader.GetAttribute("ContentType");
                    var ov = new Override(pn, ct);
                    yield return ov;
                }
                else if (!await reader.SkipContentAsync())
                    yield break;
            }
        }
#endif
        protected override void WriteImpl<TNode>(LightExcelStreamWriter writer, IEnumerable<TNode> children)
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
