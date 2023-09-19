using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using LightExcel;
using System.Diagnostics;
using System.Text;
using System.Xml.Linq;

namespace TestProject1
{
    public class M
    {
        public int Index { get; set; }
        public string? Name { get; set; }
        public DateTime Birthday { get; set; }
    }
    public class Datas
    {
        public static IEnumerable<Dictionary<string, object>> DictionarySource()
        {
            for (int i = 0; i < 10; i++)
            {
                yield return new Dictionary<string, object>
                {
                    ["Column1"] = 222,
                    ["Column2"] = 0.222,
                    ["Column3"] = 111,
                    ["Column4"] = "Hello",
                    ["Column5"] = "World",
                };
            }
        }

        public static IEnumerable<M> GetEntities()
        {
            for (int i = 0; i < 10; i++)
            {
                yield return new M
                {
                    Index = i + 1,
                    Birthday = DateTime.Now,
                    Name = "Hello"
                };
            }
        }
    }
    [TestClass]
    public class OpenXmlExcelTest
    {

        [TestMethod]
        public void CreateExcelDictionary()
        {
            var ie = Datas.DictionarySource();
            ExcelHelper excel = new ExcelHelper();
            using var trans = excel.BeginTransaction("1test.xlsx", config =>
            {
                config.AddNumberFormat("Column2");
            });
            trans.WriteExcel(ie, "sheet1");
            Process.Start("powershell", $"start {AppDomain.CurrentDomain.BaseDirectory}");
        }

        [TestMethod]
        public void CreateExcelEntity()
        {
            ExcelHelper excel = new ExcelHelper();
            excel.WriteExcel("etest.xlsx", Datas.GetEntities());
            Process.Start("powershell", $"start {AppDomain.CurrentDomain.BaseDirectory}");
        }


        [TestMethod]
        public void TemplateTest()
        {
            var ie = Datas.DictionarySource();
            ExcelHelper excel = new ExcelHelper();
            excel.WriteExcelByTemplate("12test.xlsx", "路檢報表格式.xlsx", ie, config: config =>
            {

            });
        }
    }
}