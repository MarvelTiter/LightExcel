using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;
using System.Collections;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace LightExcel.OpenXml
{
    /// <summary>
    /// xl/workbook.xml
    /// </summary>
    internal partial class SheetCollection : NodeCollectionXmlPart<Sheet>
    {
        private readonly ExcelConfiguration configuration;
        public SheetCollection(ZipArchive archive, ExcelConfiguration configuration)
            : base(archive, "xl/workbook.xml")
        {
            this.configuration = configuration;
        }

        protected override IEnumerable<Sheet> GetChildrenImpl(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.LocalName == "sheet")
                {
                    var id = reader["id", ExcelArchiveEntry.Relationships.NamespaceName] ?? throw new Exception("Excel Xml Sheet Error (without id)");
                    var name = reader["name"] ?? throw new Exception("Excel Xml Sheet Error (without name)");
                    var sid = int.Parse(reader["sheetId"] ?? throw new Exception("Excel Xml Sheet Error (without sheetId)"));
                    var sheet = new Sheet(archive!, id, name, sid);
                    sheet.NeedToSave = false;
                    yield return sheet;
                }
            }
        }

        protected override void WriteImpl(LightExcelStreamWriter writer, IEnumerable<INode> children)
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            writer.Write("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            writer.Write("<sheets>");
            foreach (INode child in children)
            {
                child.WriteToXml(writer);
            }
            writer.Write("</sheets>");
            writer.Write("</workbook>");
        }
    }

}
