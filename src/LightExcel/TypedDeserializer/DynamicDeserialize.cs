using LightExcel.Utils;

namespace LightExcel.TypedDeserializer
{
    internal static class DynamicDeserialize
    {
        internal static Func<IExcelDataReader, object> GetMapperRowDeserializer(IExcelDataReader reader, int startColumn)
        {
            var fieldCount = reader.FieldCount;

            MapperTable? table = null;

            return
                r =>
                {
                    if (table == null)
                    {
                        string[] names = new string[fieldCount];
                        for (int i = 0; i < fieldCount; i++)
                        {
                            names[i] = ReferenceHelper.ConvertX(startColumn + i);
                        }
                        table = new MapperTable(names);
                    }

                    var values = new object[fieldCount];

                    for (var iter = 0; iter < fieldCount; ++iter)
                    {
                        object obj = r.GetValue(iter);
                        values[iter] = obj;
                    }
                    return new MapperRow(table, values);
                };
        }
    }
}