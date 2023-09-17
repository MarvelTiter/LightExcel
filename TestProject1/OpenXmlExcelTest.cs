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
            using var trans = excel.BeginTransaction($"{Guid.NewGuid():N}.xlsx");
            trans.WriteExcel(ie, "sheet1");
            Process.Start("powershell", $"start {AppDomain.CurrentDomain.BaseDirectory}");
        }

        [TestMethod]
        public void ExcelCheck()
        {
            var validator = new OpenXmlValidator();
            int count = 0;
            var doc = SpreadsheetDocument.Open(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "43a8586e75114b8f8d7555be6d3ef8bb.xlsx"), true);
            StringBuilder sb = new StringBuilder();
            foreach (ValidationErrorInfo error in validator.Validate(doc))
            {
                sb.AppendLine("Error Count : " + count);
                sb.AppendLine("Description : " + error.Description);
                sb.AppendLine("Path: " + error.Path?.XPath);
                sb.AppendLine("Part: " + error.Part?.Uri);
            }
            Console.WriteLine(sb.ToString());
        }
    }
}