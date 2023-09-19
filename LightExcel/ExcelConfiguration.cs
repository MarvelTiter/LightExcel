using System.Globalization;

namespace LightExcel
{
    public class ExcelConfiguration
    {
        List<string> NeedToFormatNumberColumns = new();
        public bool UseHeader { get; set; } = true;
        public string? StartCell { get; set; }
        internal bool Readonly { get; set; }
        internal bool FillByTemplate { get; set; }
        internal int StartRowIndex { get; set; }
        /// <summary>
        /// eg. {{Name}}，false => 按顺序填充
        /// </summary>
        public bool FillWithPlacholder { get; set; } = true;
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