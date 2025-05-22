using LightExcel.OpenXml;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace LightExcel.Utils
{
    internal static class CellHelper
    {
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

        internal static Cell EmptyCell(string cr) => new() { Reference = cr, Type = "str" };
        internal static Cell CreateCell(int x, int y, object? value, ExcelColumnInfo col, bool filted, ExcelConfiguration config)
        {
            var cell = new Cell();
            cell.Reference = ReferenceHelper.ConvertXyToCellReference(x, y);
            //cell.Type = config.GetValueTypeAtRuntime && (value?.IsStringNumber() ?? false) ? "n" : ConvertCellType(col.Type ?? value?.GetType());
            //cell.Value = GetCellValue(col, value, config);
            var (v, t) = FormatCell(value, config, col);
            cell.Value = v;
            cell.Type = t;
            cell.StyleIndex = col.StyleIndex ?? (col.NumberFormat || filted ? "1" : null);
            CalcCellWidth(col, v);
            return cell;
        }
        internal static (string?, string?) FormatCell(object? value, ExcelConfiguration config, ExcelColumnInfo info)
        {
            var v = string.Empty;
            var t = "str";
            if (value == null)
            {
                return (v, t);
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
                Type? type = info.Type;
                if (info.Type == null || info.Type == typeof(object))
                {
                    type = value.GetType();
                    type = Nullable.GetUnderlyingType(type) ?? type;
                }

                if (type!.IsEnum)
                {
                    t = "str";
                    var field = type.GetField(value.ToString()!);
#if NET6_0_OR_GREATER
                    v = field?.GetCustomAttribute<System.ComponentModel.DataAnnotations.DisplayAttribute>()?.Name
                        ?? field?.GetCustomAttribute<DescriptionAttribute>()?.Description
                        ?? value.ToString();
#else
                    v = field?.GetCustomAttribute<DescriptionAttribute>()?.Description ?? value.ToString();
#endif
                }
                else if (TypeHelper.IsNumber(type))
                {
                    t = !Equals(config.CultureInfo, CultureInfo.InvariantCulture) ? "str" : "n";

                    if (type.IsAssignableFrom(typeof(decimal)))
                        v = ((decimal)value).ToString(config.CultureInfo);
                    else if (type.IsAssignableFrom(typeof(Int32)))
                        v = ((Int32)value).ToString(config.CultureInfo);
                    else if (type.IsAssignableFrom(typeof(Double)))
                        v = ((Double)value).ToString(config.CultureInfo);
                    else if (type.IsAssignableFrom(typeof(Int64)))
                        v = ((Int64)value).ToString(config.CultureInfo);
                    else if (type.IsAssignableFrom(typeof(UInt32)))
                        v = ((UInt32)value).ToString(config.CultureInfo);
                    else if (type.IsAssignableFrom(typeof(UInt16)))
                        v = ((UInt16)value).ToString(config.CultureInfo);
                    else if (type.IsAssignableFrom(typeof(UInt64)))
                        v = ((UInt64)value).ToString(config.CultureInfo);
                    else if (type.IsAssignableFrom(typeof(Int16)))
                        v = ((Int16)value).ToString(config.CultureInfo);
                    else if (type.IsAssignableFrom(typeof(Single)))
                        v = ((Single)value).ToString(config.CultureInfo);
                    else if (type.IsAssignableFrom(typeof(Single)))
                        v = ((Single)value).ToString(config.CultureInfo);
                    else
                        v = decimal.Parse(value.ToString()!).ToString(config.CultureInfo);
                }
                else if (type == typeof(bool))
                {
                    t = "b";
                    v = (bool)value ? "1" : "0";
                }
                else if (type == typeof(DateTime))
                {
                    if (!Equals(config.CultureInfo, CultureInfo.InvariantCulture))
                    {
                        t = "str";
                        v = ((DateTime)value).ToString(config.CultureInfo);
                    }
                    else if (info.Format == null)
                    {
                        t = null;
                        v = ((DateTime)value).ToOADate().ToString(CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        t = "str";
                        v = ((DateTime)value).ToString(info.Format, config.CultureInfo);
                    }
                }
                else
                {
                    v = value.ToString();
                }
            }

            return (v, t);
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

        internal static void TryGetBoolean(this Cell cell, SharedStringTable? table, out bool value)
        {
            var val = GetCellValue(cell, table);
            _ = bool.TryParse(val, out value);
        }
        internal static void TryGetDecimal(this Cell cell, SharedStringTable? table, out decimal value)
        {
            var val = GetCellValue(cell, table);
            _ = decimal.TryParse(val, out value);
        }

        internal static void TryGetDateTime(this Cell cell, SharedStringTable? table, out DateTime value)
        {
            var val = GetCellValue(cell, table);
            _ = DateTime.TryParse(val, out value);
        }
        internal static void TryGetDouble(this Cell cell, SharedStringTable? table, out double value)
        {
            var val = GetCellValue(cell, table);
            _ = double.TryParse(val, out value);
        }
        internal static void TryGetInt(this Cell cell, SharedStringTable? table, out int value)
        {
            var val = GetCellValue(cell, table);
            _ = int.TryParse(val, out value);
        }
    }
}
