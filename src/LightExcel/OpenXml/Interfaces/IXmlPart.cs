namespace LightExcel.OpenXml.Interfaces
{
    internal interface IXmlPart<T> : IDisposable where T : INode
    {
        void Write();
        void Write<TNode>(IEnumerable<TNode> children) where TNode : INode;
        void Replace<TNode>(IEnumerable<TNode> children) where TNode : INode;
    }
}
