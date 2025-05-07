using LightExcel.OpenXml.Interfaces;

namespace LightExcel.OpenXml
{
    internal class Default : INode
    {
        public Default(string? extension, string? contentType)
        {
            Extension = extension;
            ContentType = contentType;
        }

        public string? Extension { get; }
        public string? ContentType { get; }

        public void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write($"<Default Extension=\"{Extension}\" ContentType=\"{ContentType}\" />");
        }
    }
}
