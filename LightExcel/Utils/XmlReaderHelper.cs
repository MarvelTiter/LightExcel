using LightExcel.OpenXml;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace LightExcel.Utils
{
    internal static class XmlReaderHelper
    {
        public static XmlReader? GetXmlReader(this ZipArchive archive, string path)
        {
            var stream = archive.GetEntry(path)?.Open();
            if (stream == null) return null;
            return XmlReader.Create(stream);
        }

        public static LightExcelStreamWriter GetWriter(this ZipArchive archive, string path)
        {
            var stream = archive.CreateEntry(path).Open();
            return new LightExcelStreamWriter(stream, ExcelArchiveEntry.Utf8WithBom, 1024 * 512);
        }
    }
}
