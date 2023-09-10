using System.IO.Compression;
using System.Net.Mime;
using System.Text;

namespace LightExcel.OpenXml
{
    internal class ExcelDocument
    {

        public static ExcelArchiveEntry Open(string path)
        {
            var fs = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
            return new ExcelArchiveEntry(fs, zip =>
            {
                CreateZipEntry(zip, "_rels/.rels", "application/vnd.openxmlformats-package.relationships+xml", "");
                CreateZipEntry(zip, "xl/sharedStrings.xml", "application/vnd.openxmlformats-package.relationships+xml", "");
            });
        }

        private static void CreateZipEntry(ZipArchive archive, string path, string contentType, string content)
        {
            ZipArchiveEntry entry = archive.CreateEntry(path, CompressionLevel.Fastest);
            using var entryStream = entry.Open();
            using var writer = new LightExcelStreamWriter(entryStream, Utf8WithBom, 1024 * 512);
            writer.Write(content);
            
        }
    }
}
