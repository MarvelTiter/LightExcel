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
        public SharedStringTable(ZipArchive archive) : base(archive, "xl/sharedStrings.xml")
        {

        }

        protected override IEnumerable<SharedStringNode> GetChildrenImpl(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.Name == "t")
                {
                    var content = reader.ReadString();
                    var n = new SharedStringNode(content);
                    yield return n;
                }
            }
        }

        protected override void WriteImpl(LightExcelStreamWriter writer, IEnumerable<INode> children)
        {

        }

    }
}
