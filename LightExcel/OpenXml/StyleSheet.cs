using DocumentFormat.OpenXml.EMMA;
using LightExcel.OpenXml.Interfaces;
using LightExcel.OpenXml.Styles;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text.Unicode;
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

        protected override IEnumerable<INode> GetChildrenImpl(LightExcelXmlReader reader)
        {
            throw new NotImplementedException();
        }

        protected override void WriteImpl(LightExcelStreamWriter writer, IEnumerable<INode> children)
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            writer.Write("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            writer.Write("<fonts count=\"2\">");
            writer.Write("<font>");
            writer.Write("<sz val=\"11\" />");
            writer.Write("<name val=\"宋体\" />");
            writer.Write("</font>");
            writer.Write("<font>");
            writer.Write("<sz val=\"9\" />");
            writer.Write("<name val=\"宋体\" />");
            writer.Write("<family val=\"3\" />");
            writer.Write("<charset val=\"134\" />");
            writer.Write("</font>");
            writer.Write("</fonts>");
            writer.Write("<fills count=\"2\">");
            writer.Write("<fill>");
            writer.Write("<patternFill patternType=\"none\" />");
            writer.Write("</fill>");
            writer.Write("<fill>");
            writer.Write("<patternFill patternType=\"gray125\" />");
            writer.Write("</fill>");
            writer.Write("</fills>");
            writer.Write("<borders count=\"1\">");
            writer.Write("<border>");
            writer.Write("<left />");
            writer.Write("<right />");
            writer.Write("<top />");
            writer.Write("<bottom />");
            writer.Write("<diagonal />");
            writer.Write("</border>");
            writer.Write("</borders>");
            writer.Write("<cellStyleXfs count=\"1\">");
            writer.Write("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\">");
            writer.Write("<alignment vertical=\"center\" />");
            writer.Write("</xf>");
            writer.Write("</cellStyleXfs>");
            writer.Write("<cellXfs count=\"2\">");
            writer.Write("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\">");
            writer.Write("<alignment vertical=\"center\" />");
            writer.Write("</xf>");
            writer.Write("<xf numFmtId=\"10\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\">");
            writer.Write("<alignment vertical=\"center\" />");
            writer.Write("</xf>");
            writer.Write("</cellXfs>");
            writer.Write("<cellStyles count=\"1\">");
            writer.Write("<cellStyle name=\"常规\" xfId=\"0\" builtinId=\"0\" />");
            writer.Write("</cellStyles>");
            writer.Write("</styleSheet>");
        }
    }
}
