namespace LightExcel.OpenXml.Interfaces
{
    public interface INode
    {
        //string ToXmlString();
        internal void WriteToXml(LightExcelStreamWriter writer);
    }
    internal interface IStyleNode : INode
    {
        int Id { get; }
    }
}
