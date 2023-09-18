using DocumentFormat.OpenXml.Bibliography;
using LightExcel.OpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.Utils
{
    internal class CellHelper
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
            if (type.IsEnum)
            {
                var description = GetEnumDescription(type, value);
                return description ?? value.ToString();
            }
            else if (type == typeof(bool))
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
    }
}
