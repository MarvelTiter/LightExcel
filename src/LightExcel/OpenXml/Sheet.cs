using LightExcel.OpenXml.Basic;
using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;
using System.IO.Compression;
using System.Net;
using System.Xml;

namespace LightExcel.OpenXml
{
    /// <summary>
    /// xl/worksheets/sheet{id}.xml
    /// </summary>
    internal class Sheet : NodeCollectionXmlPart<Row>, INode
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

        /// <summary>
        /// 在集合中的顺序（保底机制
        /// </summary>
        public int Seq { get; set; }

        public string? Name { get; set; }
        public int SheetIdx { get; set; }
        public bool NeedToSave { get; set; } = true;
        internal override string Path => $"xl/worksheets/sheet{SheetIdx}.xml";
        public string RelPath => $"worksheets/sheet{SheetIdx}.xml";

        protected override void SetXmlReader()
        {
            reader ??= archive?.GetXmlReader($"xl/worksheets/sheet{Seq + 1}.xml");
        }

        public ExcelColumnInfo[] Columns { get; set; } = [];
        public int MaxColumnIndex { get; set; }
        public int MaxRowIndex { get; set; }

        public void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write($"""<sheet name="{Name}" sheetId="{SheetIdx}" r:id="{Id}" />""");
        }

        protected override void WriteImpl<TNode>(LightExcelStreamWriter writer, IEnumerable<TNode> children)
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            // dimension
            var dimensionWritePosition = writer.WriteAndFlush("<dimension ref=\"");
            // cols 
            var colsWritePosition = writer.WriteAndFlush("                              />");
            // <cols><col min="" max="" width="" customWidth="1"/></cols>
            var reserveLen = ReserveColsSpace(this);
            writer.Write(new string(' ', reserveLen));
            writer.Write("<sheetData>");
            foreach (var child in children)
            {
                child.WriteToXml(writer);
            }

            writer.Write("</sheetData>");
            WriteMergeInfo(writer);
            writer.WriteAndFlush("</worksheet>");
            // set dimension
            writer.SetPosition(dimensionWritePosition);
            writer.WriteAndFlush($@"A1:{ReferenceHelper.ConvertXyToCellReference(MaxColumnIndex, MaxRowIndex)}""");
            writer.SetPosition(colsWritePosition);
            CollectCols(this);
            if (ColInfos != null)
            {
                writer.WriteAndFlush(ColInfos);
            }
        }

        private void WriteMergeInfo(LightExcelStreamWriter writer)
        {
            MergeCells?.WriteToXml(writer);
        }

        internal MergeCellCollection? MergeCells { get; set; }
        internal string? ColInfos { get; set; }

        protected override IEnumerable<Row> GetChildrenImpl()
        {
            if (reader is null) yield break;
            if (!reader.IsStartWith("worksheet", XmlHelper.MainNs))
                yield break;
            if (!reader.ReadFirstContent())
                yield break;
            while (!reader.EOF)
            {
                if (reader.IsStartWith("sheetData", XmlHelper.MainNs))
                {
                    if (!reader.ReadFirstContent())
                        continue;
                    while (!reader.EOF)
                    {
                        if (reader.IsStartWith("row", XmlHelper.MainNs))
                        {
                            var row = new Row();
                            var rr = reader["r"];
                            if (rr != null)
                                row.RowIndex = int.Parse(rr);
                            if (!reader.ReadFirstContent()) continue;
                            while (!reader.EOF)
                            {
                                if (reader.IsStartWith("c", XmlHelper.MainNs))
                                {
                                    var cell = new Cell()
                                    {
                                        Reference = reader.GetAttribute("r"),
                                        Type = reader.GetAttribute("t"),
                                        StyleIndex = reader.GetAttribute("s"),
                                    };
                                    cell.Value = ReadCellValue(reader);
                                    row.AddAndFixed(cell);
                                }
                                else if (!reader.SkipContent())
                                {
                                    break;
                                }
                            }

                            yield return row;
                        }
                        else if (!reader.SkipContent())
                        {
                            break;
                        }
                    }
                }
                else if (reader.IsStartWith("cols", XmlHelper.MainNs))
                {
                    ColInfos = reader.Reader.ReadOuterXml();
                }
                else if (reader.IsStartWith("mergeCells", XmlHelper.MainNs))
                {
                    if (!reader.ReadFirstContent()) break;
                    MergeCells = new MergeCellCollection();
                    while (!reader.EOF)
                    {
                        if (reader.IsStartWith("mergeCell", XmlHelper.MainNs))
                        {
                            var merRef = reader["ref"];
                            if (merRef != null)
                                MergeCells.AppendChild(new MergeCell(merRef));
                            reader.SkipContent();
                        }
                        else if (!reader.SkipContent())
                        {
                            break;
                        }
                    }
                }
                else if (!reader.SkipContent())
                {
                    break;
                }
            }
        }

        private static string? ReadCellValue(LightExcelXmlReader reader)
        {
            string? stringValue = null;
            if (!reader.ReadFirstContent()) return stringValue;
            while (!reader.EOF)
            {
                if (reader.IsStartWith("v", XmlHelper.MainNs))
                {
                    stringValue = reader.ReadElementContentAsString();
                }
                else if (reader.IsStartWith("is", XmlHelper.MainNs))
                {
                    stringValue = reader.ReadStringContent();
                }
                else if (!reader.SkipContent())
                {
                    break;
                }
            }

            return stringValue;
        }

        private static void CollectCols(Sheet sheet)
        {
            if (!string.IsNullOrEmpty(sheet.ColInfos))
            {
                return;
            }
            var cols = sheet.Columns.Where(c => c.Width.HasValue).Select(c =>
                $"""<col min="{c.ColumnIndex}" max="{c.ColumnIndex}" width="{c.Width}" customWidth="1"/>""");
            sheet.ColInfos = $"<cols>{string.Join("", cols)}</cols>";
        }

        private static int ReserveColsSpace(Sheet sheet)
        {
            if (!string.IsNullOrEmpty(sheet.ColInfos))
            {
                return sheet.ColInfos!.Length;
            }

            // 13 {Columns}.Max(ColumnIndex).Length * {Columns}.Length * 2 + 预留宽度(10个字符=>最大宽度 9999999999)
            // <cols></cols>
            // 41
            // <col min="" max="" width="" customWidth="1"/>
            return 13 + sheet.Columns.Where(c => c.Width.HasValue || c.AutoWidth).Sum(c => 41 + StringLen(c.ColumnIndex) * 2 + 10);

            static int StringLen(int val)
            {
                int len = 1;
                while (val >= 10)
                {
                    val /= 10;
                    len++;
                }

                return len;
            }
        }
    }
}