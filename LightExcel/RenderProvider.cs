using LightExcel.Renders;
using System.Collections;
using System.Data;

namespace LightExcel
{
    public class RenderProvider
    {
        internal static IDataRender GetDataRender(Type dataType)
        {
            if (dataType == typeof(IDataReader))
            {
                //return new DataReaderRender();
            }
            else if (dataType == typeof(DataTable))
            {
                //return new DataTableRender();
            }
            else if (dataType == typeof(DataSet))
            {
                //return new DataSetRender();
            }
            else if (dataType.FindInterfaces((t, o) => t == typeof(IEnumerable), null).Length > 0)
            {
                var elementType = dataType.GetInterfaces().Where(t => IsGenericType(t)).SelectMany(t => t.GetGenericArguments()).FirstOrDefault();
                if (elementType != null)
                {
                    if (elementType == typeof(Dictionary<string, object>))
                    {
                        //return new DictionaryRender();
                    }
                    else
                    {
                        return new EnumerableEntityRender(elementType);
                    }
                }
            }
            throw new NotImplementedException($"not supported data type: {dataType}");

            bool IsGenericType(Type type1) => type1.IsGenericType && type1.GetGenericTypeDefinition() == typeof(IEnumerable<>);

        }
    }
}
