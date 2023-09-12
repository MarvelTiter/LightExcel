using DocumentFormat.OpenXml.Spreadsheet;
using LightExcel;
using System.Diagnostics;
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
            excel.WriteExcel($"{Guid.NewGuid():N}.xlsx", ie);
            Process.Start("powershell", $"start {AppDomain.CurrentDomain.BaseDirectory}");
        }

        [TestMethod]
        public void OpenTest()
        {
            var ie = Ge();
            ExcelHelper excel = new ExcelHelper();
            excel.WriteExcel(@"C:\Users\Marvel\Desktop\test\test.xlsx", ie);
        }
    }
}