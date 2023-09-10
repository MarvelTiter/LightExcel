using DocumentFormat.OpenXml.Vml.Office;
using System.IO.Compression;
using System.Text;

namespace LightExcel.OpenXml
{
    internal class ExcelArchiveEntry : IDisposable
    {
        readonly ZipArchive archive;
        private readonly Stream stream;
        private bool disposedValue;
        internal readonly static UTF8Encoding Utf8WithBom = new(true);

        public ExcelArchiveEntry(Stream stream, Action<ExcelArchiveEntry>? initital = null)
        {
            this.stream = stream;
            archive = new ZipArchive(stream, ZipArchiveMode.Update, true, Utf8WithBom);
            initital?.Invoke(this);
        }
        internal WorkBook? WorkBook { get; set; }
        internal Relationship? Relationship { get; set; }
        internal ContentTypes? ContentTypes { get; set; }

        internal void AddEntry(string path, string contentType, string content)
        {
            var zipEntry = archive.CreateEntry(path, CompressionLevel.Fastest);
            using var entryStream = zipEntry.Open();
            using var writer = new LightExcelStreamWriter(entryStream, Utf8WithBom, 1024 * 512);
            writer.Write(content);
            if (!string.IsNullOrEmpty(contentType))
                ContentTypes.Add(path, new ZipPackageInfo(entry, contentType));
        }


        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: 释放托管状态(托管对象)
                    archive?.Dispose();
                    stream?.Dispose();
                }
                disposedValue = true;
            }
        }


        public void Dispose()
        {
            // 不要更改此代码。请将清理代码放入“Dispose(bool disposing)”方法中
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
