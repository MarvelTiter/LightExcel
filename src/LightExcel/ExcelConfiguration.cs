using System.Collections.Concurrent;
using System.Globalization;

namespace LightExcel
{
    public class TransConfiguration
    {
        public bool SheetNumberFormat { get; set; }
        internal ExcelConfiguration ExcelConfig { get; set; }
        public Func<ExcelColumnInfo, bool> NumberFormatColumnFilter { get; set; }
        public TransConfiguration(ExcelConfiguration configuration)
        {
            NumberFormatColumnFilter = col => SheetNumberFormat;
            ExcelConfig = configuration;
        }

    }
    public class ExcelConfiguration
    {
        List<string> NeedToFormatNumberColumns = new();
        List<string> NeedToFormatNumberSheet = new();
        public bool UseHeader { get; set; } = true;
        public string? StartCell { get; set; }
        internal bool Readonly { get; set; }
        internal bool FillByTemplate { get; set; }
        internal int StartRowIndex { get; set; }
        public bool AutoWidth { get; set; }
        /// <summary>
        /// 例如字符串类型的数字
        /// </summary>
        public bool GetValueTypeAtRuntime { get; set; }
        /// <summary>
        /// eg. {{Name}}，false => 按顺序填充
        /// </summary>
        public bool FillWithPlacholder { get; set; } = true;
        public CultureInfo CultureInfo { get; set; } = CultureInfo.InvariantCulture;

        internal Dictionary<string, ExcelColumnInfo> DynamicColumns { get; set; } = new();
        private readonly Dictionary<int, string> nameIndex = [];
        //private Lazy<SortedSet<ExcelColumnInfo>>? _lazy;
        //internal Lazy<SortedSet<ExcelColumnInfo>> SortedColumns
        //{
        //    get
        //    {
        //        _lazy ??= new Lazy<SortedSet<ExcelColumnInfo>>(() =>
        //        {
        //            var values = DynamicColumns.Values;
        //            return new SortedSet<ExcelColumnInfo>(values, Comparer<ExcelColumnInfo>.Create((a, b) => a.ColumnIndex - b.ColumnIndex));
        //        });
        //        return _lazy;
        //    }
        //}
        public ExcelColumnInfo? this[string name] => DynamicColumns.TryGetValue(name, out ExcelColumnInfo? value) ? value : null;
        public ExcelColumnInfo? this[int index] => nameIndex.TryGetValue(index, out var n) ? this[n] : null;
        public ExcelConfiguration AddDynamicColumnInfo(string name, Action<ExcelColumnInfo> info)
        {
            if (!DynamicColumns.TryGetValue(name, out _))
            {
                ExcelColumnInfo? value = new ExcelColumnInfo(name);
                DynamicColumns.Add(name, value);
                if (value.ColumnIndex > 0)
                {
                    nameIndex[value.ColumnIndex] = name;
                }
                info.Invoke(value);
            }
            return this;
        }

        public void AddDynamicColumns(IEnumerable<ExcelColumnInfo> columns)
        {
            foreach (var item in columns)
            {
                DynamicColumns.Add(item.Name, item);
                nameIndex[item.ColumnIndex] = item.Name;
            }
        }

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
        public bool AddSheetNumberFormat { get; set; }
        public ExcelConfiguration AddNumberFormatAllSheet(string sheetName)
        {
            if (!NeedToFormatNumberSheet.Contains(sheetName))
            {
                NeedToFormatNumberSheet.Add(sheetName);
            }
            return this;
        }
        internal bool CheckCellNumberFormat(string name, string? sheetName = null)
        {
            if (!AddSheetNumberFormat)
            {
                sheetName ??= "All";
                var key = $"{sheetName}_{name}";
                return NeedToFormatNumberColumns.Contains(key);
            }
            else
            {
                return NeedToFormatNumberSheet.Contains(name) || NeedToFormatNumberSheet.Contains(sheetName);
            }
        }

    }
}