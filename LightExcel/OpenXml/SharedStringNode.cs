using LightExcel.OpenXml.Interfaces;

namespace LightExcel.OpenXml
{
    internal class SharedStringNode : INode
    {
        public SharedStringNode(string content)
        {
            Content = content;
        }

        public string Content { get; }


        public void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write($"<si><t>{Content}</t></si>");
        }
    }
}
