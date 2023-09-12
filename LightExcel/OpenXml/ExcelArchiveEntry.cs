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
            WorkBook = new WorkBook(archive, configuration);
            ContentTypes = new ContentTypes(archive);
        }
        internal WorkBook WorkBook { get; set; }
        //internal RelationshipCollection Relationship { get; set; } = new RelationshipCollection();
        internal ContentTypes ContentTypes { get; set; }

        internal void AddEntry(string path, string contentType, string content)
        {
            var zipEntry = archive.CreateEntry(path, CompressionLevel.Fastest);
            using var entryStream = zipEntry.Open();
            using var writer = new LightExcelStreamWriter(entryStream, Utf8WithBom, 1024 * 512);
            writer.Write(content);
            if (!string.IsNullOrEmpty(contentType))
                ContentTypes.AppendChild(new Override(path, contentType));
        }

        internal void LoadEntry()
        {

            ContentTypes.LoadStream("[Content_Types].xml");
            //foreach (var item in ContentTypes)
            //{
            //    Console.WriteLine(item.ToXmlString());
            //}
            WorkBook.WorkSheets.LoadStream(("xl/workbook.xml"));
            foreach (var item in WorkBook.WorkSheets)
            {
                Console.WriteLine(item.Name);
                foreach (var row in item.SheetDatas)
                {
                    Console.WriteLine($"\trow====================");
                    foreach (var cell in row.RowDatas)
                    {
                        Console.Write($"{cell.Value}|");
                    }
                    Console.WriteLine();
                }
            }
            WorkBook.Relationships.LoadStream(("xl/_rels/workbook.xml.rels"));
            //foreach (var item in WorkBook.Relationships)
            //{
            //    Console.WriteLine(item.ToXmlString());
            //}
            WorkBook.SharedStrings?.LoadStream(("xl/sharedStrings.xml"));
        }

        internal void Save()
        {
            WorkBook.Save();
            ContentTypes.Save();
        }

        private Stream? GetEntryStream(string path)
        {
            var entry = archive.GetEntry(path);
            return entry?.Open();
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
