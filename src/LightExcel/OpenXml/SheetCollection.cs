﻿using LightExcel.OpenXml.Basic;
using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;
using System.IO.Compression;

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

        protected override IEnumerable<Sheet> GetChildrenImpl()
        {
            if (reader is null) yield break;
            if (!reader.IsStartWith("workbook", XmlHelper.MainNs)) yield break;
            if (!reader.ReadFirstContent()) yield break;
            while (!reader.EOF)
            {
                if (reader.IsStartWith("sheets", XmlHelper.MainNs))
                {
                    if (!reader.ReadFirstContent())
                        continue;
                    int index = 0;
                    while (!reader.EOF)
                    {
                        if (reader.IsStartWith("sheet", XmlHelper.MainNs))
                        {
                            var id = reader["id", XmlHelper.SpreadsheetmlXmlRelationshipns] ?? throw new Exception("Excel Xml Sheet Error (without id)");
                            var name = reader["name"] ?? throw new Exception("Excel Xml Sheet Error (without name)");
                            var sid = int.Parse(reader["sheetId"] ?? throw new Exception("Excel Xml Sheet Error (without sheetId)"));
                            var sheet = new Sheet(archive!, id, name, sid)
                            {
                                NeedToSave = false,
                                Seq = index
                            };
                            index += 1;
                            // <MyNode /> 这样的节点需要调用
                            reader.SkipContent();
                            yield return sheet;
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

        protected override void WriteImpl<TNode>(LightExcelStreamWriter writer, IEnumerable<TNode> children)
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            writer.Write("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            writer.Write("<sheets>");
            foreach (var child in children)
            {
                child.WriteToXml(writer);
            }
            writer.Write("</sheets>");
            writer.Write("</workbook>");
        }
    }

}
