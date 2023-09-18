using LightExcel.OpenXml.Interfaces;
using LightExcel.OpenXml.Styles;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;

namespace LightExcel.OpenXml
{
    internal class StyleSheet : NodeCollectionXmlPart<INode>
    {
        public StyleSheet(ZipArchive archive) : base(archive, "xl/styles.xml")
        {

        }
        public FontCollection? Fonts { get; set; }
        public FillCollection? Fills { get; set; }
        public BorderCollection? Borders { get; set; }
        public NumberingFormatCollection? NumberingFormats { get; set; }
        public CellFormatCollection? CellFormats { get; set; }

        protected override IEnumerable<INode> GetChildrenImpl(XmlReader reader)
        {
            throw new NotImplementedException();
        }

        protected override void WriteImpl(LightExcelStreamWriter writer, IEnumerable<INode> children)
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            writer.Write("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            writer.Write("<fonts>");
            writer.Write("<font />");
            writer.Write("</fonts>");
            writer.Write("<fills>");
            writer.Write("<fill />");
            writer.Write("</fills>");
            writer.Write("<borders>");
            writer.Write("<border />");
            writer.Write("</borders>");
            writer.Write("<cellStyleXfs>");
            writer.Write("<xf />");
            writer.Write("</cellStyleXfs>");
            writer.Write("<cellXfs>");
            writer.Write("<xf />");
            writer.Write("<xf />");
            writer.Write("<xf />");
            writer.Write("<xf numFmtId=\"14\" applyNumberFormat=\"1\"/>");
            writer.Write("</cellXfs>");
            writer.Write("</styleSheet>");
        }
    }
}
