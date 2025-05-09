using System.Collections;
using LightExcel.OpenXml.Interfaces;

namespace LightExcel.OpenXml.Basic;

internal abstract class SimpleNodeCollectionXmlPart<T> : INodeCollection<T>, INode where T : INode
{
    public IEnumerator<T> GetEnumerator() => Children.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
    public List<T> Children { get; } = [];
    public void AppendChild(T child) => Children.Add(child);
    public int Count => Children.Count;
    public virtual void WriteToXml(LightExcelStreamWriter writer)
    {
        
    }
}