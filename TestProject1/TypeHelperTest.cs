using LightExcel.Utils;

namespace TestProject1
{
    [TestClass]
    public class TypeHelperTest
    {
        [TestMethod]
        public void IsNumberType()
        {
            var val1 = TypeHelper.IsNumber(typeof(int?));
            Assert.IsTrue(val1);

            var val2 = TypeHelper.IsNumber(typeof(int));
            Assert.IsTrue(val2);
        }
    }
}