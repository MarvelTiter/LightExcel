using LightExcel.Utils;

namespace TestProject1
{
    [TestClass]
    public class CellRefConvertTest
    {
        [TestMethod]
        public void XyToRef()
        {
            var r1 = ReferenceHelper.ConvertXyToCellReference(1, 1);
            Assert.IsTrue(r1 == "A1");
            var r2 = ReferenceHelper.ConvertXyToCellReference(2, 2);
            Assert.IsTrue(r2 == "B2");
            var r3 = ReferenceHelper.ConvertXyToCellReference(27, 3);
            Assert.IsTrue(r3 == "AA3");
            var r4 = ReferenceHelper.ConvertXyToCellReference(55, 5);
            Assert.IsTrue(r4 == "BC5");
        }

        [TestMethod]
        public void RefToXy()
        {
            var r1 = ReferenceHelper.ConvertCellReferenceToXY("A1");
            Assert.IsTrue(r1.X == 1 && r1.Y == 1);
            var r2 = ReferenceHelper.ConvertCellReferenceToXY("B2");
            Assert.IsTrue(r2.X == 2 && r2.Y == 2);
            var r3 = ReferenceHelper.ConvertCellReferenceToXY("AA3");
            Assert.IsTrue(r3.X == 27 && r3.Y == 3);
            var r4 = ReferenceHelper.ConvertCellReferenceToXY("BC5");
            Assert.IsTrue(r4.X == 55 && r4.Y == 5);
        }
    }
}