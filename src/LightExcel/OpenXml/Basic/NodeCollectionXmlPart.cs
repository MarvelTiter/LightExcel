using System.IO.Compression;
using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;
using System.Collections;
namespace LightExcel.OpenXml.Basic;

internal abstract class NodeCollectionXmlPart<T>(ZipArchive archive, string path) : XmlPart<T>(archive, path), INodeCollection<T> where T : INode
{
    public List<T> Children { get; } = new List<T>();

    public int Count => Children.Count;

    public void AppendChild(T child) => Children.Add(child);

    public override void Write()
    {
        using var writer = archive!.GetWriter(Path);
        WriteImpl(writer, Children);
    }
    protected virtual IEnumerable<T> GetChildren()
    {
        SetXmlReader();
        if (reader == null) yield break;
        foreach (var item in GetChildrenImpl())
        {
            yield return item;
        }
    }

    protected abstract IEnumerable<T> GetChildrenImpl();

    public virtual IEnumerator<T> GetEnumerator() => GetChildren().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}