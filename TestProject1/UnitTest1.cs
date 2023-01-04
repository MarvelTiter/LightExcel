using LightExcel;

namespace TestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var ie = Ge();
            RenderProvider.GetDataRender(ie.GetType());
        }

        class Test01
        {
            public int Prop1 { get; set; }
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