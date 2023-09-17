using System.IO.Compression;
using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;

namespace LightExcel.OpenXml
{
    internal abstract class NodeCollectionXmlPart<T> : XmlPart<T>, INodeCollection<T> where T : INode
    {

        public NodeCollectionXmlPart(ZipArchive archive, string path) : base(archive, path)
        {

        }

        public int Count => cached?.Count ?? GetChildren().Count();

        public void AppendChild(T child)
        {
            cached ??= new List<T>();
            cached.Add(child);
        }
        public override void Write()
        {
            if (cached == null) return;
            using var writer = archive!.GetWriter(Path);
            WriteImpl(writer, cached.Cast<INode>());
        }
    }
}
