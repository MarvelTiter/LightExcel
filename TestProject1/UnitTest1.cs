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
            [ExcelColumn(Name = "����1")]
            public int Prop1 { get; set; }
            [ExcelColumn(Name = "����2")]
            public int Prop2 { get; set; }
        }

        IEnumerable<Test01> Ge()
        {
            for (int i = 0; i < 10; i++)
            {
                yield return new Test01();
            }
        }
    }
}