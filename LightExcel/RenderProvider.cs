using LightExcel.Renders;
using System.Collections;
using System.Data;

namespace LightExcel
{
    public class RenderProvider
    {
        internal static IDataRender GetDataRender(Type dataType, ExcelConfiguration configuration)
        {
            if (typeof(IDataReader).IsAssignableFrom(dataType))
            {
                return new DataReaderRender(configuration);
            }
            else if (dataType == typeof(DataTable))
            {
                return new DataTableRender(configuration);
            }
            else if (dataType.FindInterfaces((t, o) => t == typeof(IEnumerable), null).Length > 0)
            {
                var elementType = dataType.GetInterfaces().Where(t => IsGenericType(t)).SelectMany(t => t.GetGenericArguments()).FirstOrDefault();
                if (elementType != null)
                {
                    if (elementType == typeof(Dictionary<string, object>))
                    {
                        return new DictionaryRender(configuration);
                    }
                    else
                    {
                        return new EnumerableEntityRender(elementType, configuration);
                    }
                }
            }
            throw new NotImplementedException($"not supported data type: {dataType}");

            bool IsGenericType(Type type1) => type1.IsGenericType && type1.GetGenericTypeDefinition() == typeof(IEnumerable<>);

        }
    }
}
