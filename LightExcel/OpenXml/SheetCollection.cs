using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace LightExcel.OpenXml
{
    /// <summary>
    /// xl/worksheets/sheet{id}.xml
    /// </summary>
    internal class Sheet : XmlPart
    {
        public Sheet(ZipArchive archive, string id, string name, int sid)
        {
            this.archive = archive;
            Id = id;
            Name = name;
            SheetIdx = sid;
            SheetDatas = new RowCollection(archive, Path);
        }
        public Sheet(ZipArchive archive, string name, int index)
        {
            this.archive = archive;
            Name = name;
            SheetIdx = index;
            SheetDatas = new RowCollection(archive, Path);
        }

        public string Id { get; set; } = $"R{Guid.NewGuid():N}";
        public string? Name { get; set; }
        public int SheetIdx { get; set; }
        public bool NeedToSave { get; set; } = true;
        public string Path => $"xl/worksheets/sheet{SheetIdx}.xml";
        public string RelPath => $"worksheets/sheet{SheetIdx}.xml";
        public RowCollection SheetDatas { get; set; }

        public void AppendRow(Row row)
        {
            SheetDatas.Append(row);
        }

        internal override void Save()
        {
            writer!.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            writer!.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
        }
    }
    /// <summary>
    /// xl/workbook.xml
    /// </summary>
    internal partial class SheetCollection : XmlPart<Sheet>
    {
        private readonly ExcelHelperConfiguration configuration;
        public SheetCollection(ZipArchive archive, ExcelHelperConfiguration configuration) : base(archive)
        {
            this.configuration = configuration;
        }

        internal Sheet? this[string id]
        {
            get
            {
                return Children!.FirstOrDefault(s => s.Id == id);
            }
        }

        protected override IEnumerable<Sheet> GetChildren()
        {
            if (reader == null)
            {
                yield break;
            }
            while (reader.Read())
            {
                if (reader.LocalName == "sheet")
                {
                    var id = reader["id", ExcelArchiveEntry.Relationships.NamespaceName] ?? throw new Exception("Excel Xml Sheet Error (without id)");
                    var name = reader["name"] ?? throw new Exception("Excel Xml Sheet Error (without name)");
                    var sid = int.Parse(reader["sheetId"] ?? throw new Exception("Excel Xml Sheet Error (without sheetId)"));
                    var sheet = new Sheet(archive!, id, name, sid);
                    sheet.NeedToSave = false;
                    Children.Add(sheet);
                    yield return sheet;
                }
            }
        }



        internal override void Save()
        {
            throw new NotImplementedException();
        }
    }

}
