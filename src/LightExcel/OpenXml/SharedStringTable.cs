using LightExcel.OpenXml.Basic;
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
        //private readonly IDictionary<int, string> values = new Dictionary<int, string>();

        public SharedStringTable(ZipArchive archive) : base(archive, "xl/sharedStrings.xml")
        {
            Flush();
        }

        internal string? this[int index]
        {
            get
            {
                if (index < 0 || index >= Children.Count)
                    return null;
                return Children[index].Content;
            }
        }

        private void Flush()
        {
            // _ = GetChildren().ToList();
            Children.Clear();
            foreach (var s in GetChildren())
            {
                Children.Add(s);
            }
        }

        public override IEnumerator<SharedStringNode> GetEnumerator() => Children.GetEnumerator();

        protected override IEnumerable<SharedStringNode> GetChildrenImpl()
        {
            if (reader is null) yield break;
            if (!reader.IsStartWith("sst", XmlHelper.MainNs)) yield break;
            if (!reader.ReadFirstContent()) yield break;
            _ = int.TryParse(reader["count"], out var count);
            _ = int.TryParse(reader["uniqueCount"], out var uniqueCount);
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

        protected override void WriteImpl<TNode>(LightExcelStreamWriter writer, IEnumerable<TNode> children)
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