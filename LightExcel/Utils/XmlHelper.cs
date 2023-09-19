using LightExcel.OpenXml;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace LightExcel.Utils
{
    internal static class XmlHelper
    {
        public static LightExcelXmlReader? GetXmlReader(this ZipArchive archive, string path)
        {
            var stream = archive.GetEntry(path)?.Open();
            if (stream == null) return null;
            return new LightExcelXmlReader(stream);
        }

        public static LightExcelStreamWriter GetWriter(this ZipArchive archive, string path)
        {
            var stream = archive.CreateEntry(path).Open();
            return new LightExcelStreamWriter(stream, ExcelArchiveEntry.Utf8WithBom, 1024 * 512);
        }

        public const string SpreadsheetmlXmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        public const string SpreadsheetmlXmlStrictns = "http://purl.oclc.org/ooxml/spreadsheetml/main";
        public const string SpreadsheetmlXmlRelationshipns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public const string SpreadsheetmlXmlStrictRelationshipns = "http://purl.oclc.org/ooxml/officeDocument/relationships";
        public const string ContentTypesXmlns = "http://schemas.openxmlformats.org/package/2006/content-types";

        public readonly static string[] MainNs = new[] { SpreadsheetmlXmlns, SpreadsheetmlXmlStrictns };
        public readonly static string[] RelaNs = new[] { SpreadsheetmlXmlRelationshipns, SpreadsheetmlXmlStrictRelationshipns };

        public static string ReadStringContent(this LightExcelXmlReader reader)
        {
            var result = new StringBuilder();
            if (!reader.ReadFirstContent())
                return string.Empty;

            while (!reader.EOF)
            {
                if (reader.IsStartWith("t", XmlHelper.MainNs))
                {
                    result.Append(reader.ReadElementContentAsString());
                }
                else if (reader.IsStartWith("r", XmlHelper.MainNs))
                {
                    var runs = ReadRichTextRun(reader);
                    foreach (var r in runs)
                    {
                        result.Append(r);
                    }
                }
                else if (!reader.SkipContent())
                {
                    break;
                }
            }

            return result.ToString();
        }
        private static IEnumerable<string> ReadRichTextRun(LightExcelXmlReader reader)
        {
            if (!reader.ReadFirstContent())
                yield break;

            while (!reader.EOF)
            {
                if (reader.IsStartWith("t", XmlHelper.MainNs))
                {
                    yield return reader.ReadElementContentAsString();
                }
                else if (!reader.SkipContent())
                {
                    yield break;
                }
            }
        }
    }
}
