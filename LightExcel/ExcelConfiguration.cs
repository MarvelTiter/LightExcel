using DocumentFormat.OpenXml.Packaging;
using LightExcel.CellSetting;
using LightExcel.Enums;

namespace LightExcel
{
    public class ExcelConfiguration
    {
        public string SheetName { get; set; } = "sheet";
        public string? TemplatePath { get; set; }
        public bool AllowAppendSheet { get; set; } = true;

        Dictionary<string, StyleHelper> formatters = new();

        public void AddNumberFormat(string column, NumberFormat format)
        {
            StyleHelper setting = GetHelper(column);
            setting.AddStyle(new CellNumberFormat(format));
        }

        public void AddNumberFormat<TProperty>(string column, NumberFormat format, Func<TProperty, bool> filter)
        {
            StyleHelper setting = GetHelper(column);
            setting.AddStyle(new CellNumberFormat(format, obj =>
            {
                if (obj == null) return false;
                return filter.Invoke((TProperty)obj);
            }));
        }

        internal void AddStyle(string column, IExcelCellStyle style)
        {
            StyleHelper setting = GetHelper(column);
            setting.AddStyle(style);
        }

        public void AddFillStyle(string column, string background, string foreground)
        {
            AddFillStyle(column, FillType.Solid, background, foreground);
        }


        public void AddFillStyle(string column, FillType type, string background, string foreground)
        {
            StyleHelper setting = GetHelper(column);
            setting.AddStyle(new CellFillStyle(type, background, foreground));
        }


        public void AddFillStyle<TProperty>(string column, string background, string foreground, Func<TProperty, bool> filter)
        {
            AddFillStyle(column, FillType.Solid, background, foreground, filter);
        }

        public void AddFillStyle<TProperty>(string column, FillType type, string background, string foreground, Func<TProperty, bool> filter)
        {
            StyleHelper setting = GetHelper(column);
            setting.AddStyle(new CellFillStyle(type, background, foreground, obj =>
            {
                if (obj == null) return false;
                return filter.Invoke((TProperty)obj);
            }));
        }

        private StyleHelper GetHelper(string column)
        {
            if (!formatters.TryGetValue(column, out var setting))
            {
                setting = new StyleHelper();
                formatters[column] = setting;
            }
            return setting;
        }


        internal bool HasStyle(string key, object? value)
        {
            if (!formatters.TryGetValue(key, out var setting))
            {
                return false;
            }
            return setting.HasStyle(value);
        }

        internal uint? GetStyleIndex(string key, WorkbookPart workbookPart)
        {
            if (!formatters.TryGetValue(key, out var setting))
            {
                return null;
            }
            return setting.GetStyleIndex(workbookPart);
        }
    }
}