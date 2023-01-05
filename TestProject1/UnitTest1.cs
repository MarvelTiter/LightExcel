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
            [ExcelColumn(Name = " Ù–‘1")]
            public int Prop1 { get; set; }
            [ExcelColumn(Name = " Ù–‘2")]
            public int Prop2 { get; set; }
        }

        IEnumerable<Dictionary<string, object>> Ge()
        {
            for (int i = 0; i < 10; i++)
            {
                yield return new Dictionary<string, object>
                {
                    ["Column1"] = 222,
                    ["Column2"] = "≤‚ ‘",
                    ["Column3"] = 111,
                    ["Column4"] = "Hello",
                    ["Column5"] = "World",

                };
            }
        }
    }
}