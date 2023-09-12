using System.Collections;
using System.IO.Compression;
using System.Text;

namespace LightExcel.OpenXml
{
    internal class RelationshipCollection : XmlPart<Relationship>
    {
        public RelationshipCollection(ZipArchive archive) : base(archive)
        {

        }
        protected override IEnumerable<Relationship> GetChildren()
        {
            if (reader == null)
            {
                yield break;
            }
            while (reader.Read())
            {
                if (reader.LocalName == "Relationship")
                {
                    var id = reader["Id"] ?? throw new Exception("Excel Xml Relationship Error (without id)");
                    var type = reader["Type"] ?? throw new Exception("Excel Xml Relationship Error (without type)");
                    var target = reader["Target"] ?? throw new Exception("Excel Xml Relationship Error (without target)");
                    var rel = new Relationship(id, type, target);
                    Children.Add(rel);
                    yield return rel;
                }
            }
        }

        internal override void Save()
        {
            throw new NotImplementedException();
        }
    }

    internal class Relationship : Node
    {
        const string TYPE_PREFIX = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
        public string Id { get; set; }
        /// <summary>
        /// worksheet / sharedStrings / styles
        /// </summary>
        public string Type { get; set; }
        public string Target { get; set; }
        public Relationship(string id, string type, string target)
        {
            Id = id;
            Type = type;
            Target = target;
        }
        internal override string ToXmlString()
        {
            return $"<Relationship Id=\"{Id}\" Type=\"{TYPE_PREFIX}{Type}\" Target=\"{Target}\" />";
        }
    }
}
