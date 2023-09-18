namespace LightExcel.OpenXml.Interfaces
{
    internal interface INode
    {
        //string ToXmlString();
        void WriteToXml(LightExcelStreamWriter writer);
    }
    internal interface IStyleNode : INode
    {
        int Id { get; }
    }
}
