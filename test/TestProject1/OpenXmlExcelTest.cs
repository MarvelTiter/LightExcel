using LightExcel;
using System.Diagnostics;
using System.Text;
using System.Xml.Linq;

namespace TestProject1
{
    [TestClass]
    public class OpenXmlExcelTest
    {
        [TestMethod]
        public void CreateExcelDictionary()
        {
            var ie = Datas.DictionarySource();
            ExcelHelper excel = new ExcelHelper();
            using var trans = excel.BeginTransaction("dic-test.xlsx", config =>
            {
                config.AddNumberFormat("Column2");
                config.AutoWidth = true;
            });
            trans.WriteExcel(ie);
            Process.Start("powershell", $"start {AppDomain.CurrentDomain.BaseDirectory}");
        }

        [TestMethod]
        public async Task CreateExcelDictionaryAsync()
        {
            var ie = Datas.DictionarySourceAsync();
            ExcelHelper excel = new ExcelHelper();

            await excel.WriteExcelAsync("async-dic-test.xlsx", ie, config: config =>
            {
                config.AddNumberFormat("Column2");
                config.AutoWidth = true;
            });
            Process.Start("powershell", $"start {AppDomain.CurrentDomain.BaseDirectory}");
        }

        [TestMethod]
        public void CreateExcelEntity()
        {
            ExcelHelper excel = new ExcelHelper();
            excel.WriteExcel("entity-test.xlsx", Datas.GetEntities().ToList(), config: c =>
            {
                //c.AutoWidth = true;
            });
            Process.Start("powershell", $"start {AppDomain.CurrentDomain.BaseDirectory}");
        }

        [TestMethod]
        public async Task CreateExcelEntityAsync()
        {
            ExcelHelper excel = new ExcelHelper();
            await excel.WriteExcelAsync("async-entity-test.xlsx", Datas.GetEntitiesAsync(), config: c =>
             {
                 c.AutoWidth = true;
             });
            Process.Start("powershell", $"start {AppDomain.CurrentDomain.BaseDirectory}");
        }

        [TestMethod]
        public void CreateExcelDataTable()
        {
            ExcelHelper excel = new();
            excel.WriteExcel("dt-test.xlsx", Datas.GetDataTable());
            Process.Start("powershell", $"start {AppDomain.CurrentDomain.BaseDirectory}");
        }

        [TestMethod]
        public void TemplateTest()
        {
            var ie = Datas.DictionarySource();
            ExcelHelper excel = new ExcelHelper();
            excel.WriteExcelByTemplate("template-test.xlsx", "template.xlsx", ie, config: config =>
            {
                //config.FillWithPlacholder = true;
            });
            Process.Start("powershell", $"start {AppDomain.CurrentDomain.BaseDirectory}");
        }

        [TestMethod]
        public async Task TemplateTestAsync()
        {
            var ie = Datas.DictionarySourceAsync();
            ExcelHelper excel = new ExcelHelper();
            await excel.WriteExcelByTemplateAsync("async-template-test.xlsx", "template.xlsx", ie, config: config =>
             {
                 //config.FillWithPlacholder = true;
             });
            Process.Start("powershell", $"start {AppDomain.CurrentDomain.BaseDirectory}");
        }
    }
}