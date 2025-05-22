namespace LightExcel.OpenXml.Interfaces
{
    internal interface INodeCollection<T> : IEnumerable<T> where T : INode
    {
        List<T> Children { get; }
        void AppendChild(T child);
        int Count { get; }
    }
}
