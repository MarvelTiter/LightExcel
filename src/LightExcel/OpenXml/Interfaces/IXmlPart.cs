namespace LightExcel.OpenXml.Interfaces
{
    internal interface IXmlPart<T> : IDisposable, IEnumerable<T> where T : INode
    {
        void Write();
        void Write(IEnumerable<INode> children);
        void Replace(IEnumerable<INode> children);
    }
}
