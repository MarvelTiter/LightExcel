using LightExcel.OpenXml;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace LightExcel.Utils;

internal static class CellHelper
{
    private static readonly Dictionary<Type, Func<object, CultureInfo, string>> numberFormatters = new()
    {
        [typeof(decimal)] = (v, culture) => ((decimal)v).ToString(culture),
        [typeof(double)] = (v, culture) => ((double)v).ToString(culture),
        [typeof(int)] = (v, culture) => ((int)v).ToString(culture),
        [typeof(long)] = (v, culture) => ((long)v).ToString(culture),
        [typeof(short)] = (v, culture) => ((short)v).ToString(culture),
        [typeof(uint)] = (v, culture) => ((uint)v).ToString(culture),
        [typeof(ulong)] = (v, culture) => ((ulong)v).ToString(culture),
        [typeof(ushort)] = (v, culture) => ((ushort)v).ToString(culture),
        [typeof(float)] = (v, culture) => ((float)v).ToString(culture)
    };
    public static bool IsNumeric(this string self)
    {
        var match = Regex.IsMatch(self, @"([1-9]\d*\.?\d*)|(0\.\d*[1-9])");
        return match;
    }
    internal static bool IsStringNumber(this object? value)
    {
        return value?.ToString()?.IsNumeric() ?? false;
    }
    internal static string ConvertCellType(Type? type, bool shared = false)
    {
        if (type == null) return "str";
        if (TypeHelper.IsNumber(type)) return "n";
        type = Nullable.GetUnderlyingType(type) ?? type;
        return Type.GetTypeCode(type) switch
        {
            TypeCode.Boolean => "b",
            //TypeCode.DateTime => "d",
            TypeCode.String when shared => "s",
            _ => "str"
        };
    }

    internal static Cell EmptyCell(string cr) => new(null, cr, "str");
    internal static Cell CreateCell(int x, int y, object? value, ExcelColumnInfo col, TransConfiguration transConfig)
    {
        //var cell = new Cell();
        //cell.Reference = ReferenceHelper.ConvertXyToCellReference(x, y);
        //cell.Type = config.GetValueTypeAtRuntime && (value?.IsStringNumber() ?? false) ? "n" : ConvertCellType(col.Type ?? value?.GetType());
        //cell.Value = GetCellValue(col, value, config);
        var cell = FormatCell(ReferenceHelper.ConvertXyToCellReference(x, y), value, transConfig, col);
        //cell.Value = v;
        //cell.Type = t;
        CalcCellWidth(col, cell.Value);
        return cell;
    }
    internal static Cell FormatCell(string r, object? value, TransConfiguration transConfig, ExcelColumnInfo info)
    {
        var v = string.Empty;
        var t = "str";
        string? s = null;
        var config = transConfig.ExcelConfig;
        if (value == null)
        {
            return new Cell(null, r, null);
        }
        if (value is string str)
        {
            v = value.ToString();
        }
        else if (info.Format != null && value is IFormattable formattable)
        {
            v = formattable.ToString(info.Format, config.CultureInfo);
        }
        else
        {
            (v, t, s) = FormatByType(value, info, config);
        }
        //cell.Value = v;
        //cell.Type = t;
        var filted = transConfig.NumberFormatColumnFilter(info);
        s ??= info.StyleIndex ?? (info.NumberFormat || filted ? "1" : null);
        return new Cell(v, r, t, s);

        static bool IsNumberValue(object value, ExcelConfiguration config, out string v)
        {
            if (value is decimal dec)
                v = dec.ToString(config.CultureInfo);
            else if (value is int i32)
                v = i32.ToString(config.CultureInfo);
            else if (value is double dou)
                v = dou.ToString(config.CultureInfo);
            else if (value is long i64)
                v = i64.ToString(config.CultureInfo);
            else if (value is uint u32)
                v = u32.ToString(config.CultureInfo);
            else if (value is ushort u16)
                v = u16.ToString(config.CultureInfo);
            else if (value is ulong u64)
                v = u64.ToString(config.CultureInfo);
            else if (value is short i16)
                v = i16.ToString(config.CultureInfo);
            else if (value is float f)
                v = f.ToString(config.CultureInfo);
            else
            {
                var s = value.ToString();
                if (decimal.TryParse(s, out var ddd))
                {
                    v = ddd.ToString(config.CultureInfo);
                }
                else
                {
                    v = string.Empty;
                    return false;
                }
            }
            return true;
        }

        static (string value, string? type, string? style) FormatByType(object value, ExcelColumnInfo info, ExcelConfiguration config)
        {
            var type = info.Type ?? value.GetType();
            type = Nullable.GetUnderlyingType(type) ?? type;

            // 使用switch表达式优化类型判断
            return type switch
            {
                Type enumType when enumType.IsEnum => FormatEnum(value, enumType),
                Type numType when TypeHelper.IsNumber(numType) => FormatNumber(value, numType, config),
                Type boolType when boolType == typeof(bool) => FormatBoolean((bool)value),
                Type dateType when dateType == typeof(DateTime) => FormatDateTime((DateTime)value, info, config),
                _ => FormatFallback(value, config, info)
            };
        }

        static (string value, string? type, string? style) FormatEnum(object value, Type enumType)
        {
            var field = enumType.GetField(value.ToString()!);
            if (field == null)
                return (value.ToString()!, "str", null);

#if NET6_0_OR_GREATER
            var displayName = field.GetCustomAttribute<System.ComponentModel.DataAnnotations.DisplayAttribute>()?.Name;
            if (!string.IsNullOrEmpty(displayName))
                return (displayName, "str", null);
#endif

            var description = field.GetCustomAttribute<DescriptionAttribute>()?.Description;
            return (description ?? value.ToString()!, "str", null);
        }
        static (string value, string? type, string? style) FormatNumber(object value, Type numberType, ExcelConfiguration config)
        {
            var useStringType = !Equals(config.CultureInfo, CultureInfo.InvariantCulture);
            var culture = config.CultureInfo;

            if (numberFormatters.TryGetValue(numberType, out var formatter))
            {
                return (formatter(value, config.CultureInfo), useStringType ? "str" : "n", null);
            }
            // 回退处理
            var stringValue = value.ToString();
            if (decimal.TryParse(stringValue, out var decimalValue))
                return (decimalValue.ToString(culture), useStringType ? "str" : "n", null);

            return (stringValue ?? string.Empty, "str", null);
        }
        static (string value, string? type, string? style) FormatBoolean(bool value) => (value ? "1" : "0", "b", null);
        static (string value, string? type, string? style) FormatDateTime(DateTime value, ExcelColumnInfo info, ExcelConfiguration config)
        {
            if (!Equals(config.CultureInfo, CultureInfo.InvariantCulture))
            {
                return (value.ToString(config.CultureInfo), "str", null);
            }

            if (info.Format == null)
            {
                return (value.ToOADate().ToString(CultureInfo.InvariantCulture), null, info.StyleIndex ?? "2");
            }

            return (value.ToString(info.Format, config.CultureInfo), "str", null);
        }
        static (string value, string? type, string? style) FormatFallback(object value, ExcelConfiguration config, ExcelColumnInfo info)
        {

            if (IsNumberValue(value, config, out var s))
            {
                var useStringType = !Equals(config.CultureInfo, CultureInfo.InvariantCulture);
                return (s, useStringType ? "str" : "n", null);
            }

            if (value is bool b)
                return FormatBoolean(b);

            if (value is DateTime dt)
                return FormatDateTime(dt, info, config);

            return (value.ToString() ?? string.Empty, "str", null);
        }
    }

    internal static void CalcCellWidth(ExcelColumnInfo col, string? value)
    {
        if (!col.AutoWidth)
        {
            return;
        }
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }
        if (!col.Width.HasValue)
        {
            col.Width = CalcStringWidth(col.Name);
        }
        double cellLength = CalcStringWidth(value!);
        col.Width = Math.Max(col.Width.Value, cellLength);

    }
    internal static double CalcStringWidth(string value)
    {
        var mg = Regex.Matches(value, "[^\x00-\xff]+");
        int dl = 0;
        int sl = 0;
        foreach (Match m in mg)
        {
            if (m.Success)
            {
                dl += m.Value.Length;
            }
        }
        sl = value!.Length - dl;

        var cellLength = dl * 2.1 + sl * 1.1;
        return Math.Round(cellLength, 3);
    }


    // private static string? GetEnumDescription(Type type, object value)
    // {
    //     var field = type.GetField(value.ToString()!)!;
    //     var descAttr = field.GetCustomAttribute<DescriptionAttribute>();
    //     return descAttr?.Description;
    // }

    internal static string? GetCellValue(this Cell cell, SharedStringTable? table)
    {
        // if (cell == null) return null;
        if (cell.Type == "s")
        {
            if (int.TryParse(cell.Value, out var s))
            {
                return table?[s];
            }
        }
        return cell.Value;
    }

    internal static void TryGetBoolean(this Cell cell
        , SharedStringTable? table
        , int index
        , ExcelConfiguration configuration
        , out bool value)
    {
        var val = GetCellValue(cell, table);
        _ = bool.TryParse(val, out value);
    }
    internal static void TryGetDecimal(this Cell cell
        , SharedStringTable? table
        , int index
        , ExcelConfiguration configuration
        , out decimal value)
    {
        var val = GetCellValue(cell, table);
        _ = decimal.TryParse(val, out value);
    }

    internal static void TryGetDateTime(this Cell cell
        , SharedStringTable? table
        , int index
        , ExcelConfiguration configuration
        , out DateTime value)
    {
        var val = GetCellValue(cell, table);
        if (!DateTime.TryParse(val, out value))
        {
            // ExcelColumnInfo的Index是从1开始的，参数的index是从0开始
            var ci = configuration[index + 1];
            if (ci?.Format is not null)
            {
                _ = DateTime.TryParseExact(val, ci.Format, CultureInfo.InvariantCulture, DateTimeStyles.None, out value);
            }
            else if (double.TryParse(val, out var oad))
            {
                value = DateTime.FromOADate(oad);
            }
        }
    }
    internal static void TryGetDouble(this Cell cell
        , SharedStringTable? table
        , int index
        , ExcelConfiguration configuration
        , out double value)
    {
        var val = GetCellValue(cell, table);
        _ = double.TryParse(val, out value);
    }
    internal static void TryGetInt(this Cell cell
        , SharedStringTable? table
        , int index
        , ExcelConfiguration configuration
        , out int value)
    {
        var val = GetCellValue(cell, table);
        _ = int.TryParse(val, out value);
    }
}
