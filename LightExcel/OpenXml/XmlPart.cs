using System.Collections;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace LightExcel.OpenXml
{
    internal abstract class XmlPart : IDisposable
    {
        protected Stream? Stream;
        protected XmlReader? reader;
        private bool disposedValue;
        protected ZipArchive? archive;
        protected LightExcelStreamWriter? writer;
        internal virtual void Save() { }

        internal virtual void LoadStream(string path)
        {
            if (archive == null) { return; }
            Stream = archive.GetEntry(path)?.Open();
            if (Stream == null)
            {
                Stream = archive.CreateEntry(path).Open();
                writer = new LightExcelStreamWriter(Stream, ExcelArchiveEntry.Utf8WithBom, 512 * 1024);
            }
            else
            {
                reader = XmlReader.Create(Stream);
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    Stream?.Dispose();
                    Stream = null;
                    reader?.Dispose();
                    reader = null;
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

    internal abstract class XmlPart<T> : XmlPart, IEnumerable<T>
    {
        internal IList<T> Children { get; set; } = new List<T>();
        protected abstract IEnumerable<T> GetChildren();
        public XmlPart(ZipArchive archive)
        {
            this.archive = archive;
        }
        internal override void LoadStream(string path)
        {
            base.LoadStream(path);
        }

        public void Append(T child)
        {
            Children.Add(child);
        }

        public IEnumerator<T> GetEnumerator()
        {
            if (Children == null) return GetEnumerator();
            else
            {
                return Children!.GetEnumerator();
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
