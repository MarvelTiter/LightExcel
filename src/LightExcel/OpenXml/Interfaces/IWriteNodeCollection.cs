namespace LightExcel.OpenXml.Interfaces
{
    internal interface IWriteNodeCollection<T> where T : INode
    {
        void Write(LightExcelStreamWriter writer, IEnumerable<INode> children);
    }
}
