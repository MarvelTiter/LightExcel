using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace LightExcel.OpenXml
{
    internal class StyleSheet
    {
        public object? Fonts { get; set; }
        public object? Fills { get; set; }
        public object? Borders { get; set; }
        public object? NumberingFormats { get; set; }
        public object? CellFormats { get; set; }
    }
}
