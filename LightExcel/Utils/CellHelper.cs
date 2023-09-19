using LightExcel.OpenXml;
using System.ComponentModel;
using System.Reflection;

namespace LightExcel.Utils
{
    internal static class CellHelper
    {
        internal static string ConvertCellType(Type? type, bool shared = false)
        {
            if (type == null) return "str";
            if (TypeHelper.IsNumber(type)) return "n";
            type = Nullable.GetUnderlyingType(type) ?? type;
            return Type.GetTypeCode(type) switch
            {
                TypeCode.Boolean => "b",
                TypeCode.DateTime => "d",
                TypeCode.String when shared => "s",
                _ => "str"
            };
        }

        internal static Cell EmptyCell(string cr) => new() { Reference = cr, Type = "str" };

        internal static string? GetCellValue(ExcelColumnInfo col, object? value, ExcelConfiguration configuration)
        {
            if (value == null) return null;
            var type = value.GetType();
            var underType = Nullable.GetUnderlyingType(type) ?? type;
            if (underType.IsEnum)
            {
                var description = GetEnumDescription(type, value);
                return description ?? value.ToString();
            }
            else if (underType == typeof(bool))
            {
                return (bool)value ? "1" : "0";
            }
            else if (value is IFormattable formattable)
            {
                if (col.Format != null)
                    return formattable.ToString(col.Format, configuration.CultureInfo);
                else
                    return formattable.ToString();
            }
            else
            {
                return value.ToString();
            }
        }


        private static string? GetEnumDescription(Type type, object value)
        {
            var field = type.GetField(value.ToString()!)!;
            var descAttr = field.GetCustomAttribute<DescriptionAttribute>();
            return descAttr?.Description;
        }

        internal static string? GetCellValue(this Cell cell, SharedStringTable? table)
        {
            if (cell == null) return null;
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
