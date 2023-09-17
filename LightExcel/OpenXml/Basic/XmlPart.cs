using System.Collections;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;

namespace LightExcel.OpenXml
{
    internal abstract partial class XmlPart<T> : IXmlPart<T> where T : INode
    {
        protected Stream? stream;
        private bool disposedValue;
        protected ZipArchive? archive;
        internal virtual string Path { get; }
        public XmlPart(ZipArchive archive, string path)
        {
            this.archive = archive;
            Path = path;
        }

        protected IList<T>? cached;

        public virtual void Write()
        {

        }

        public void Write(IEnumerable<INode> children)
        {
            using var writer = archive!.GetWriter(Path);
            WriteImpl(writer, children);
        }

        protected abstract void WriteImpl(LightExcelStreamWriter writer, IEnumerable<INode> children);

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    archive?.Dispose();
                    archive = null;
                }
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }


    }
}
