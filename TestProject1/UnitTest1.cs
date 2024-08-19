using LightExcel;
using LightExcel.Attributes;
using System.Diagnostics;
using System.Text;

namespace TestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var ie = Ge();
            ExcelHelper excel = new ExcelHelper();
            using var ms = new MemoryStream();
            excel.WriteExcel(ms, ie, config: c =>
            {
                c.AutoWidth = true;
                c.AddDynamicColumnInfo("Column1", col =>
                {
                    col.Width = 20;
                }).AddDynamicColumnInfo("Column3", col =>
                {
                    col.AutoWidth = false;
                });
            });
            File.WriteAllBytes($"{Guid.NewGuid():N}.xlsx", ms.ToArray());
            Process.Start("powershell", $"start {AppDomain.CurrentDomain.BaseDirectory}");
        }

        class Test01
        {
            [ExcelColumn(Name = "属性1")]
            public int Prop1 { get; set; }
            [ExcelColumn(Name = "属性2")]
            public int Prop2 { get; set; }
        }

        IEnumerable<Dictionary<string, object>> Ge()
        {
            for (int i = 0; i < 10; i++)
            {
                yield return new Dictionary<string, object>
                {
                    ["Column1"] = 222,
                    ["Column2"] = new string('测', (i + 1) * 2),
                    ["Column3"] = 111,
                    ["Column4"] = new string('A', (i + 1) * 2),
                    ["Column5"] = "World",

                };
            }
        }

        [TestMethod]
        public void TestRead()
        {
            ExcelHelper excel = new ExcelHelper();
            var reader = excel.ReadExcel(@"E:\Documents\Desktop\导出流水\整备质量数据表 (润通).xlsx");
            while (reader.NextResult())
            {
                Console.WriteLine($"================={reader.CurrentSheetName}================");
                while (reader.Read())
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        Console.Write($"{reader[i]}, ");
                    }
                    Console.Write(Environment.NewLine);
                }
                break;
            }
        }

        [TestMethod]
        public void TemplateTest()
        {
            ExcelHelper excel = new ExcelHelper();
            const string Path1 = "templateTest.xlsx";
            if (File.Exists(Path1))
                File.Delete(Path1);
            using var ms = new MemoryStream();
            using var template = File.Open("template.xlsx", FileMode.Open, FileAccess.Read, FileShare.Read);
            excel.WriteExcelByTemplate(ms, template, Ge());
            File.WriteAllBytes(Path1, ms.ToArray());
            Process.Start("powershell", $"start {AppDomain.CurrentDomain.BaseDirectory}");
        }

        [TestMethod]
        public void EmbeddedTemplateTest()
        {
            ExcelHelper excel = new ExcelHelper();
            const string Path1 = "templateEmbededTest.xlsx";
            if (File.Exists(Path1))
                File.Delete(Path1);
            var files = GetType().Assembly.GetManifestResourceNames().FirstOrDefault(s => s.EndsWith("template.xlsx"));
            using var ms = new MemoryStream();
            using var template = GetType().Assembly.GetManifestResourceStream(files!)!;
            excel.WriteExcelByTemplate(ms, template, Ge());
            File.WriteAllBytes(Path1, ms.ToArray());
            Process.Start("powershell", $"start {AppDomain.CurrentDomain.BaseDirectory}");
        }

    }
}