using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using LightExcel;
using System.Diagnostics;
using System.Text;
using System.Xml.Linq;

namespace TestProject1
{
    [TestClass]
    public class OpenXmlExcelTest
    {
        IEnumerable<Dictionary<string, object>> Ge()
        {
            for (int i = 0; i < 10; i++)
            {
                yield return new Dictionary<string, object>
                {
                    ["Column1"] = 222,
                    ["Column2"] = "测试",
                    ["Column3"] = 111,
                    ["Column4"] = "Hello",
                    ["Column5"] = "World",

                };
            }
        }
        [TestMethod]
        public void CreateExcel()
        {
            var ie = Ge();
            ExcelHelper excel = new ExcelHelper();
            using var trans = excel.BeginTransaction("1test.xlsx", config =>
            {
            });
            trans.WriteExcel(ie, "sheet1");
            Process.Start("powershell", $"start {AppDomain.CurrentDomain.BaseDirectory}");
        }

       
        [TestMethod]
        public void TemplateTest()
        {
            var ie = Ge();
            ExcelHelper excel = new ExcelHelper();
            excel.WriteExcelByTemplate("12test.xlsx", "1test.xlsx", ie);

        }
    }
}