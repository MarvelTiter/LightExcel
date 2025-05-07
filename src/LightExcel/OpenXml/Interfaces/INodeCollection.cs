namespace LightExcel.OpenXml.Interfaces
{
    internal interface INodeCollection<T> where T : INode
    {
        void AppendChild(T child);
        int Count { get; }
    }
}
