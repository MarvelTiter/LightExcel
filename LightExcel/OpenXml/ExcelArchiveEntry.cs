using DocumentFormat.OpenXml.Vml.Office;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace LightExcel.OpenXml
{
    internal class ExcelArchiveEntry : IDisposable
    {
        readonly ZipArchive archive;
        private readonly Stream stream;
        private readonly ExcelHelperConfiguration configuration;
        private bool disposedValue;
        internal readonly static UTF8Encoding Utf8WithBom = new(true);
        internal readonly static XNamespace Main_Xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        internal readonly static XNamespace Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        public ExcelArchiveEntry(Stream stream, ExcelHelperConfiguration configuration)
        {
            this.stream = stream;
            this.configuration = configuration;
            archive = new ZipArchive(stream, ZipArchiveMode.Update, true, Utf8WithBom);
            ContentTypes = new ContentTypes(archive);
        }
        internal WorkBook? WorkBook { get; set; }
        //internal RelationshipCollection Relationship { get; set; } = new RelationshipCollection();
        internal ContentTypes ContentTypes { get; set; }
        internal void AddWorkBook()
        {
            WorkBook = new WorkBook(archive, this, configuration);
            ContentTypes.AppendChild(new Override("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"));
        }
        internal void AddEntry(string path, string contentType, string content)
        {
            var zipEntry = archive.CreateEntry(path, CompressionLevel.Fastest);
            using var entryStream = zipEntry.Open();
            using var writer = new LightExcelStreamWriter(entryStream, Utf8WithBom, 1024 * 512);
            writer.Write(content);
            if (!string.IsNullOrEmpty(contentType))
                ContentTypes.AppendChild(new Override(path, contentType));
        }

        internal void SetTemplate(Stream templateStream)
        {
            templateStream.Seek(0, SeekOrigin.Begin);
            templateStream.CopyTo(stream);
            templateStream.Dispose();
        }

        internal void LoadEntry()
        {
        }

        internal void Save()
        {
            WorkBook?.Save();
            ContentTypes.Write();
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
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
