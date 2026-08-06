using System.IO.Compression;
using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;
using System.Collections;
using System.Runtime.CompilerServices;
namespace LightExcel.OpenXml.Basic;

internal abstract class NodeCollectionXmlPart<T>(ZipArchive archive, string path, ExcelConfiguration configuration)
    : XmlPart<T>(archive, path, configuration), INodeCollection<T> where T : INode
{
    public List<T> Children { get; } = new List<T>();

    public virtual int Count => Children.Count;

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

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        Children.Clear();
    }
#if NET8_0_OR_GREATER
    protected virtual async IAsyncEnumerable<T> GetChildrenAsync([EnumeratorCancellation] CancellationToken cancellationToken)
    {
        SetXmlReader();
        if (reader == null) yield break;
        await foreach (var item in GetChildrenImplAsync(cancellationToken))
        {
            yield return item;
        }
    }
    protected abstract IAsyncEnumerable<T> GetChildrenImplAsync(CancellationToken cancellationToken);
    public virtual IAsyncEnumerator<T> GetAsyncEnumerator(CancellationToken cancellationToken = default)
        => GetChildrenAsync(cancellationToken).GetAsyncEnumerator(cancellationToken);
#endif
}