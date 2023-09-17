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
                return cached?[index]?.Content;
            }
        }

        protected override IEnumerable<SharedStringNode> GetChildrenImpl(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.Name == "sst")
                {
                    int.TryParse(reader["count"], out var count);
                    int.TryParse(reader["uniqueCount"], out var uniqueCount);
                    RefCount = count;
                    UniqueCount = uniqueCount;
                }
                else if (reader.Name == "t")
                {
                    var content = reader.ReadString();
                    var n = new SharedStringNode(content);
                    yield return n;
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
