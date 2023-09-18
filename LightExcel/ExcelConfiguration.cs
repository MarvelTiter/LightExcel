using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using LightExcel.Utils;
using System.Globalization;

namespace LightExcel
{
    public class ExcelConfiguration
    {
        List<string> NeedToFormatNumberColumns = new();
        public bool UseHeader { get; set; } = true;
        public string StartCell { get; set; } = "A1";
        internal bool Readonly { get; set; }
        public CultureInfo CultureInfo { get; set; } = CultureInfo.InvariantCulture;
        public ExcelConfiguration AddNumberFormat(string name, string? sheetName = null)
        {
            sheetName ??= "All";
            var key = $"{sheetName}_{name}";
            if (!NeedToFormatNumberColumns.Contains(key))
            {
                NeedToFormatNumberColumns.Add(key);
            }
            return this;
        }
        internal bool CheckCellNumberFormat(string name, string? sheetName = null)
        {
            sheetName ??= "All";
            var key = $"{sheetName}_{name}";
            return NeedToFormatNumberColumns.Contains(key);
        }

    }
}