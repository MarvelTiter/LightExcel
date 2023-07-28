using LightExcel;
using LightExcel.Attributes;
using System.Collections;

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
            excel.WriteExcel("test.xlsx", ie);
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
                    ["Column2"] = "测试",
                    ["Column3"] = 111,
                    ["Column4"] = "Hello",
                    ["Column5"] = "World",

                };
            }
        }

        [TestMethod]
        public void TestRead()
        {
            ExcelHelper excel = new ExcelHelper();
            var reader = excel.ReadExcel(@"E:\Statistics\2023 06月 每月统计.xlsx");
            reader.NextResult();
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
            var ie = Ge();
            excel.WriteExcel("E:\\Documents\\Downloads\\test1.xlsx", @"E:\Documents\Downloads\路檢報表格式.xlsx", ie);
            excel.WriteExcel("E:\\Documents\\Downloads\\test2.xlsx", @"E:\Documents\Downloads\路檢報表格式.xlsx", ie);
        }
    }
}