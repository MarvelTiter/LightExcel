using LightExcel.OpenXml.Interfaces;

namespace LightExcel.OpenXml
{
    internal class Relationship : INode
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

        public void WriteToXml(LightExcelStreamWriter writer)
        {
            writer!.Write($"<Relationship Id=\"{Id}\" Type=\"{TYPE_PREFIX}{Type}\" Target=\"{Target}\" />");
        }
    }
}
