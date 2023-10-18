using LightExcel;

namespace TestProject1
{
    [TestClass]
    public class OpenXmlExcelReadTest
    {
        [TestMethod]
        public void ExcelReaderTest()
        {
            ExcelHelper excel = new ExcelHelper();
            using var reader = excel.ReadExcel("1test.xlsx", config: config =>
            {
                config.StartCell = "A1";
            });
            while (reader.NextResult())
            {
                Console.WriteLine(reader.CurrentSheetName);
                while (reader.Read())
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        Console.Write($"{reader[i]}\t");
                    }
                    Console.WriteLine();
                }
            }
        }

        [TestMethod]
        public void ExcelReaderTestEntity()
        {
            ExcelHelper excel = new ExcelHelper();
            var result = excel.QueryExcel<M>("etest.xlsx");
        }
        [TestMethod]
        public void ExcelReaderTestDynamic()
        {
            ExcelHelper excel = new ExcelHelper();
            var reader = excel.ReadExcel("C:\\Users\\Marvel\\Desktop\\截止20231017二期车证.xlsx");
            while (reader.NextResult())
            {
                while (reader.Read())
                {
                    Console.WriteLine($"D: {reader["车牌号码"]}, E: {reader["车辆识别代码后4位"]}");
                }
            }
            //var resule = excel.QueryExcel("C:\\Users\\Marvel\\Desktop\\截止20231017二期车证.xlsx");
            //foreach (var field in result)
            //{
            //    Console.WriteLine($"D: {field.D}, E: {field.E}");
            //}
        }

    }
}