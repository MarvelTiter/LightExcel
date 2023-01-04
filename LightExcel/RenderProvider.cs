using System.Collections;
using System.Data;

namespace LightExcel
{
    public class RenderProvider
    {
        public static IDataRender GetDataRender(Type dataType)
        {               
            if (dataType == typeof(IDataReader))
            {
                return new DataReaderRender();
            }
            else if (dataType == typeof(DataTable))
            {
                return new DataTableRender();
            }
            else if (dataType == typeof(DataSet))
            {
                return new DataSetRender();
            }
            else if (dataType.FindInterfaces((t, o) => t == typeof(IEnumerable), null).Length > 0)
            {
                var elementType = dataType.GetInterfaces().Where(t => IsGenericType(t)).SelectMany(t => t.GetGenericArguments()).ToList();
                return new EnumerableEntityRender(elementType.First());
            }
            throw new NotImplementedException($"not supported data type: {dataType}");

            bool IsGenericType(Type type1) => type1.IsGenericType && type1.GetGenericTypeDefinition() == typeof(IEnumerable<>);

        }
    }
}
