namespace LightExcel.OpenXml.Interfaces
{
    internal interface INodeCollection<T>
        : IEnumerable<T>
#if NET8_0_OR_GREATER
        , IAsyncEnumerable<T>
#endif
        where T : INode
    {
        List<T> Children { get; }
        void AppendChild(T child);
        int Count { get; }
    }
}
