using LightExcel.OpenXml.Interfaces;

namespace LightExcel.OpenXml.Styles
{
    internal class CellFormatCollection
    {

    }

    internal class CellFormat : IStyleNode
    {
        private int? fontId;
        private int? borderId;
        private int? fillId;
        private int? numberFormatId;

        public int Id { get; }
        public int? NumberFormatId { get => numberFormatId; set { numberFormatId = value; ApplyNumberFormat = true; } }
        public int? FontId { get => fontId; set { fontId = value; ApplyFont = true; } }
        public int? FillId { get => fillId; set { fillId = value; ApplyFill = true; } }
        public int? BorderId { get => borderId; set { borderId = value; ApplyBorder = true; } }
        public bool ApplyNumberFormat { get; set; }
        public bool ApplyFill { get; set; }
        public bool ApplyFont { get; set; }
        public bool ApplyBorder { get; set; }

        private string ApplyNumberString => ApplyNumberFormat ? "applyNumberFormat=\"1\"" : "";
        private string ApplyFillString => ApplyFill ? "" : "";
        private string ApplyFontString => ApplyFont ? "" : "";
        private string ApplyBorderString => ApplyBorder ? "" : "";

        public void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write("<xf >");
        }
    }
}
