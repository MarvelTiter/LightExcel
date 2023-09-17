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
        private static readonly string defaultSharedString = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"></sst>";
        public static ExcelArchiveEntry Open(string path, ExcelHelperConfiguration configuration)
        {
            var fs = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
            var zip = new ExcelArchiveEntry(fs, configuration);
            //TODO: 打开操作
            zip.LoadEntry();
            return zip;
        }

        public static ExcelArchiveEntry Create(string path, ExcelHelperConfiguration configuration)
        {
            //TODO: buffer size of create file
            var fs = File.Create(path, 1024 * 512);
            var zip = new ExcelArchiveEntry(fs, configuration);
            //TODO: 创建初始化操作
            zip.AddEntry("_rels/.rels", "application/vnd.openxmlformats-package.relationships+xml", defaultRels);
            zip.AddEntry("xl/sharedStrings.xml", "application/vnd.openxmlformats-package.relationships+xml", defaultSharedString);
            zip.AddWorkBook();
            zip.WorkBook!.AddSharedStringTable();
            //zip.WorkBook.AddStyleSheet();
            return zip;
        }
    }
}
