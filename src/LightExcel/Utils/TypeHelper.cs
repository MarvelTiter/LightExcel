using LightExcel.Attributes;
using System.Reflection;

namespace LightExcel.Utils
{
    internal static class TypeHelper

    {
        public static bool IsNumber(Type type)
        {
            type = Nullable.GetUnderlyingType(type) ?? type;
            return Type.GetTypeCode(type) switch
            {
                TypeCode.UInt16 or TypeCode.UInt32 or TypeCode.UInt64 or TypeCode.Int16 or TypeCode.Int32 or TypeCode.Int64 or TypeCode.Decimal or TypeCode.Double or TypeCode.Single => true,
                _ => false,
            };
        }

        public static IEnumerable<ExcelColumnInfo> CollectEntityInfo(this Type type, Action<ExcelColumnInfo>? callback = null)
        {
            var properties = type.GetProperties();
            int index = 1;
            foreach (var prop in properties)
            {
                ExcelColumnAttribute? excelColumnAttribute = prop.GetCustomAttribute<ExcelColumnAttribute>();
                if (excelColumnAttribute?.Ignore ?? false) continue;
#if NET6_0_OR_GREATER
                var displayAttribute = prop.GetCustomAttribute<System.ComponentModel.DataAnnotations.DisplayAttribute>();
                var col = new ExcelColumnInfo(excelColumnAttribute?.Name ?? displayAttribute?.Name ?? prop.Name);
#else
                var col = new ExcelColumnInfo(excelColumnAttribute?.Name ??  prop.Name);
#endif

                col.Ignore = excelColumnAttribute?.Ignore ?? false;
                col.Property = new Property(prop);
                col.Type = prop.PropertyType;
                col.NumberFormat = excelColumnAttribute?.NumberFormat ?? false;
                col.Format = excelColumnAttribute?.Format;
                col.ColumnIndex = index++;
                col.AutoWidth = excelColumnAttribute?.AutoWidth ?? false;
                col.Width = excelColumnAttribute?.Width;
                callback?.Invoke(col);
                yield return col;
            }
        }
    }
}
