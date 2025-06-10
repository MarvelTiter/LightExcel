using System.Data;

namespace TestProject1
{
    public class Datas
    {
        public static IEnumerable<Dictionary<string, object>> DictionarySource()
        {
            for (int i = 0; i < 10; i++)
            {
                yield return new Dictionary<string, object>
                {
                    ["Column0"] = i,
                    ["Column1"] = DateTime.Now,
                    ["Column2"] = 0.222,
                    ["Column3"] = 111,
                    ["Column4"] = "Hello",
                    ["Column5"] = "World",
                };
            }
        }

        public static async IAsyncEnumerable<Dictionary<string, object?>> DictionarySourceAsync()
        {
            for (int i = 0; i < 10; i++)
            {
                yield return new Dictionary<string, object?>
                {
                    ["Column0"] = i,
                    ["Column1"] = DateTime.Now,
                    ["Column2"] = 0.222,
                    ["Column3"] = 111,
                    ["Column4"] = "Hello",
                    ["Column5"] = "World",
                };
                await Task.Delay(5);
            }
        }

        public static IEnumerable<Model> GetEntities()
        {
            for (int i = 0; i < 10; i++)
            {
                yield return new Model
                {
                    Index = i + 1,
                    Birthday = DateTime.Now.Date,
                    Name = "Hello",
                    Birthday2 = DateTime.Now
                };
            }
        }

        public static async IAsyncEnumerable<Model> GetEntitiesAsync()
        {
            for (int i = 0; i < 10; i++)
            {
                yield return new Model
                {
                    Index = i + 1,
                    Birthday = DateTime.Now,
                    Name = "Hello",
                    Birthday2 = DateTime.Now
                };
                await Task.Delay(5);
            }
        }

        public static DataTable GetDataTable()
        {
            DataTable dt = new DataTable();
            for (int i = 0; i < 10; i++)
            {
                dt.Columns.Add($"Column{i}");
            }
            for (int i = 0; i < 10; i++)
            {
                var row = dt.NewRow();
                for (int j = 0; j < 10; j++)
                {
                    //row[j] = $"D({i}x{j})";
                    row[j] = $"{i * j}";
                }
                dt.Rows.Add(row);
            }
            return dt;
        }
    }
}