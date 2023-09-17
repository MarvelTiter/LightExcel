using System.IO.Compression;
using System.Net.Mime;
using System.Text;

namespace LightExcel.OpenXml
{
    internal class ExcelDocument
    {
        //TODO: 初始化优化
        private static readonly string defaultRels =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml"" Id=""Rfc2254092b6248a9"" />
</Relationships>";
        public static ExcelArchiveEntry Open(string path, ExcelHelperConfiguration configuration)
        {
            var fs = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
            var zip = new ExcelArchiveEntry(fs, configuration);
            //TODO: 打开操作
            //zip.WorkBook.InitStyleSheet();
            zip.WorkBook.InitSharedStringTable();
            return zip;
        }

        public static ExcelArchiveEntry Create(string path, ExcelHelperConfiguration configuration)
        {
            //TODO: buffer size of create file
            var fs = File.Create(path, 1024 * 512);
            var zip = new ExcelArchiveEntry(fs, configuration);
            //TODO: 创建初始化操作
            zip.AddEntry("_rels/.rels", "application/vnd.openxmlformats-package.relationships+xml", defaultRels);
            zip.AddWorkBook();
            zip.WorkBook.AddSharedStringTable();
            //zip.WorkBook.AddStyleSheet();
            return zip;
        }

        public static ExcelArchiveEntry CreateByTemplate(string path, string template, ExcelHelperConfiguration configuration)
        {
            using var templateStream = File.Open(template, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
            var fs = File.Create(path, 1024 * 512);
            templateStream.CopyTo(fs);
            var zip = new ExcelArchiveEntry(fs, configuration);
            //zip.SetTemplate(templateStream);
            zip.WorkBook.InitSharedStringTable();
            return zip;
        }
    }
}
