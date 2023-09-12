using LightExcel.OpenXml.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace LightExcel.OpenXml
{
    internal class StyleSheet
    {
        public FontCollection? Fonts { get; set; }
        public FillCollection? Fills { get; set; }
        public BorderCollection? Borders { get; set; }
        public NumberingFormatCollection? NumberingFormats { get; set; }
        public CellFormatCollection? CellFormats { get; set; }
    }
}
