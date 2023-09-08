using DocumentFormat.OpenXml.Packaging;
using LightExcel.Renders;
using System.Collections;
using System.Data;

namespace LightExcel
{
    public class RenderProvider
    {
        public static IDataRender GetDataRender(Type dataType, WorkbookPart workbookPart, ExcelConfiguration configuration)
        {
            if (dataType == typeof(IDataReader))
            {
                return new DataReaderRender(workbookPart, configuration);
            }
            else if (dataType == typeof(DataTable))
            {
                return new DataTableRender(workbookPart, configuration);
            }
            else if (dataType.FindInterfaces((t, o) => t == typeof(IEnumerable), null).Length > 0)
            {
                var elementType = dataType.GetInterfaces().Where(t => IsGenericType(t)).SelectMany(t => t.GetGenericArguments()).FirstOrDefault();
                if (elementType != null)
                {
                    if (elementType == typeof(Dictionary<string, object>))
                    {
                        return new DictionaryRender(workbookPart, configuration);
                    }
                    else
                    {
                        return new EnumerableEntityRender(elementType, workbookPart, configuration);
                    }
                }
            }
            throw new NotImplementedException($"not supported data type: {dataType}");

            bool IsGenericType(Type type1) => type1.IsGenericType && type1.GetGenericTypeDefinition() == typeof(IEnumerable<>);

        }
    }
}
