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

		internal ConcurrentDictionary<string, ExcelColumnInfo> DynamicColumns { get; set; } = new();

		public ExcelColumnInfo? this[string name] => DynamicColumns.TryGetValue(name, out ExcelColumnInfo? value) ? value : null;

		public ExcelConfiguration AddDynamicColumnInfo(string name, Action<ExcelColumnInfo> info)
		{
			var columnInfo = DynamicColumns.GetOrAdd(name, n =>
				{
					return new ExcelColumnInfo(n);
				});
			info.Invoke(columnInfo);
			return this;
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