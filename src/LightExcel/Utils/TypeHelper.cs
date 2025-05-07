namespace LightExcel.Utils
{
#if DEBUG
    public class TypeHelper
#else
    internal class TypeHelper
#endif
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
    }
}
