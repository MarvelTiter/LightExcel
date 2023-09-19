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
    }
}