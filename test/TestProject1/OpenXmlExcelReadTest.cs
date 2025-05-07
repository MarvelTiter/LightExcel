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
            var result = excel.QueryExcel<M>("etest.xlsx","Sheet1");
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
        // "C:\Users\Marvel\Desktop\驾驶人证件过期短信提醒\20250416模板\1驾驶人临近期满换证（期满日期前3个月）_结果.xlsx"
        [TestMethod]
        public void ExcelReaderTestDynamic2()
        {
            ExcelHelper excel = new ExcelHelper();
            var reader = excel.ReadExcel("C:\\Users\\Marvel\\Desktop\\驾驶人证件过期短信提醒\\20250416模板\\1驾驶人临近期满换证（期满日期前3个月）_结果.xlsx");
            while (reader.NextResult())
            {
                while (reader.Read())
                {
                    //Console.WriteLine($"Index: {reader.RowIndex}, F: {reader[5]}, H: {reader[7]}");
                    _ = $"Index: {reader.RowIndex}, F: {reader[5]}, H: {reader[7]}";
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