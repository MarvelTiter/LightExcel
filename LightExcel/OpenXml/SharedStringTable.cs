using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;
using System.IO.Compression;
using System.Xml;

namespace LightExcel.OpenXml
{
    /// <summary>
    /// xl/sharedStrings.xml
    /// </summary>
    internal class SharedStringTable : NodeCollectionXmlPart<SharedStringNode>
    {
        public int RefCount { get; set; }
        public int UniqueCount { get; set; }
        public SharedStringTable(ZipArchive archive) : base(archive, "xl/sharedStrings.xml")
        {

        }

        internal string? this[int index]
        {
            get
            {
                if (cached == null)
                {
                    return GetChildren().ElementAt(index).Content;
                }
                else
                {
                    return cached[index].Content;
                }
            }
        }

        protected override IEnumerable<SharedStringNode> GetChildrenImpl(LightExcelXmlReader reader)
        {
            if (!reader.IsStartWith("sst", XmlHelper.MainNs)) yield break;
            if (!reader.ReadFirstContent()) yield break;
            int.TryParse(reader["count"], out var count);
            int.TryParse(reader["uniqueCount"], out var uniqueCount);
            RefCount = count;
            UniqueCount = uniqueCount;
            while (!reader.EOF)
            {
                if (reader.IsStartWith("si", XmlHelper.MainNs))
                {
                    var content = reader.ReadStringContent();
                    yield return new SharedStringNode(content);
                }
                else if (!reader.SkipContent())
                {
                    break;
                }
            }
        }

        protected override void WriteImpl(LightExcelStreamWriter writer, IEnumerable<INode> children)
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            writer.Write($"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{RefCount}\" uniqueCount=\"{UniqueCount}\">");
            foreach (var child in children)
            {
                child.WriteToXml(writer);
            }
            writer.Write("</sst>");
        }
    }
}
