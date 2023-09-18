using System.Globalization;

namespace LightExcel
{
    public class ExcelHelperConfiguration
    {
        public bool UseHeader { get; set; } = true;
        public string StartCell { get; set; } = "A1";
        internal bool Readonly { get; set; }
        public CultureInfo CultureInfo { get; set; } = CultureInfo.InvariantCulture;
    }
}