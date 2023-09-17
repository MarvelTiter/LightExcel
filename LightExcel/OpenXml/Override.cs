using LightExcel.OpenXml.Interfaces;

namespace LightExcel.OpenXml
{
    internal class Override : INode
    {
        public Override(string? PartName, string? ContentType)
        {
            this.PartName = PartName;
            this.ContentType = ContentType;
        }

        public string? PartName { get; }
        public string? ContentType { get; }

        public void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write($"<Override PartName=\"{PartName}\" ContentType=\"{ContentType}\" />");
        }
    }
}
