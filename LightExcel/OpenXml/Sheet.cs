using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;
using System.IO.Compression;
using System.Xml;

namespace LightExcel.OpenXml
{
    /// <summary>
    /// xl/worksheets/sheet{id}.xml
    /// </summary>
    internal class Sheet : XmlPart<Row>, INode
    {
        public Sheet(ZipArchive archive, string id, string name, int sid) : base(archive, "")
        {
            // read
            Id = id;
            Name = name;
            SheetIdx = sid;
        }
        public Sheet(ZipArchive archive, string name, int index) : base(archive, "")
        {
            Id = $"R{Guid.NewGuid():N}";
            Name = name;
            SheetIdx = index;
        }

        public string Id { get; set; }
        public string? Name { get; set; }
        public int SheetIdx { get; set; }
        public bool NeedToSave { get; set; } = true;
        internal override string Path => $"xl/worksheets/sheet{SheetIdx}.xml";
        public string RelPath => $"worksheets/sheet{SheetIdx}.xml";

        public int MaxColumnIndex { get; set; }
        public int MaxRowIndex { get; set; }

        public void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write($@"<sheet name=""{Name}"" sheetId=""{SheetIdx}"" r:id=""{Id}"" />");
        }

        protected override void WriteImpl(LightExcelStreamWriter writer, IEnumerable<INode> children)
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            var dimensionWritePosition = writer.WriteAndFlush("<dimension ref=\"");
            writer.Write("                              />");
            writer.Write("<sheetData>");
            // dimension
            foreach (INode child in children)
            {
                child.WriteToXml(writer);
            }

            writer.Write("</sheetData>");
            writer.WriteAndFlush("</worksheet>");
            // set dimension
            writer.SetPosition(dimensionWritePosition);
            writer.WriteAndFlush($@"A1:{ReferenceHelper.ConvertXyToCellReference(MaxColumnIndex, MaxRowIndex)}""");
        }

        protected override IEnumerable<Row> GetChildrenImpl(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.Name == "row")
                {
                    var row = new Row();
                    while (reader.Read())
                    {
                        if (reader.Name == "c")
                        {
                            var c = new Cell
                            {
                                Reference = reader["r"],
                                Type = reader["t"],
                                StyleIndex = reader["s"],
                            };
                            reader.Read();
                            c.Value = reader.ReadInnerXml();
                            row.RowDatas.Add(c);
                            break;
                        }
                    }
                    yield return row;
                }
            }
        }
    }

}
