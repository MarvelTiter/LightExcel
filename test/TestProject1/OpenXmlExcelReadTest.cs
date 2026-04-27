using LightExcel;

namespace TestProject1
{
    public class Tj22Xfsjfx
    {
        /// <summary>
        /// 来源渠道
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "来源渠道")]
        public string? LYQD { get; set; }
        /// <summary>
        /// 序号
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "序号")]
        public int? XH { get; set; }
        /// <summary>
        /// 日期
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "日期")]
        public DateTime? RQ { get; set; }
        /// <summary>
        /// 姓名
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "姓名")]
        public string? XM { get; set; }
        /// <summary>
        /// 电话
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "电话")]
        public string? DH { get; set; }
        /// <summary>
        /// 车牌
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "车牌")]
        public string? CP { get; set; }
        /// <summary>
        /// 身份证号
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "身份证号")]
        public string? SFZH { get; set; }
        /// <summary>
        /// 工单内容
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "工单内容")]
        public string? GDNR { get; set; }
        /// <summary>
        /// 反映情况
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "反映情况")]
        public string? FYQK { get; set; }
        /// <summary>
        /// 分类
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "分类")]
        public string? FL { get; set; }
        /// <summary>
        /// 各科分所
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "各科分所")]
        public string? GKFS { get; set; }
        /// <summary>
        /// 工单类型
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "工单类型")]
        public string? GDLX { get; set; }
        /// <summary>
        /// 业务分类
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "业务分类")]
        public string? YWFL { get; set; }
        /// <summary>
        /// 处理结果
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "处理结果")]
        public string? CLJG { get; set; }
        /// <summary>
        /// 回复人
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "回复人")]
        public string? HFR { get; set; }
        /// <summary>
        /// 岗位
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "岗位")]
        public string? GW { get; set; }
        /// <summary>
        /// 民/辅警/邮政人员
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "民/辅警/邮政人员")]
        public string? MFJYZRY { get; set; }
        /// <summary>
        /// 回复方式
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "回复方式")]
        public string? HFFS { get; set; }
        /// <summary>
        /// 是否满意
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "是否满意")]
        public string? SFMY { get; set; }
        /// <summary>
        /// 转派
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "转派")]
        public string? ZP { get; set; }
        /// <summary>
        /// 预留1
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "预留1")]
        public string? YL1 { get; set; }
        /// <summary>
        /// 预留2
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "预留2")]
        public string? YL2 { get; set; }
        /// <summary>
        /// 预留3
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "预留3")]
        public string? YL3 { get; set; }
        /// <summary>
        /// 预留4
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "预留4")]
        public string? YL4 { get; set; }
        /// <summary>
        /// 预留5
        /// </summary>
        [LightExcel.Attributes.ExcelColumn(Name = "预留5")]
        public string? YL5 { get; set; }
    }
    [TestClass]
    public class OpenXmlExcelReadTest
    {
        [TestMethod]
        public void ExcelReaderTest2()
        {
            ExcelHelper excel = new ExcelHelper();
            var reader = excel.QueryExcel<Tj22Xfsjfx>("C:\\Users\\Marvel\\Desktop\\tttt\\数字看板数据统计\\xfmb.xlsx");
            foreach (var item in reader)
            {
            }
        }

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
            var result = excel.QueryExcel<Model>("entity-test.xlsx", config: config =>
            {

            });
            foreach (var item in result)
            {
                Console.WriteLine($"{item.Name} - {item.Birthday} - {item.Birthday2}");
            }
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
            int actionCount = 0;
            int totalRowCount = 0;
            int notMapCount = 0;
            var reader = excel.ReadExcel("C:\\Users\\Marvel\\Desktop\\驾驶人证件过期短信提醒\\20250416模板\\1驾驶人临近期满换证（期满日期前3个月）_结果.xlsx");
            while (reader.NextResult())
            {
                while (reader.Read())
                {
                    var kx = reader.GetValue(5)?.ToString();
                    var yz = reader.GetValue(7)?.ToString();
                    //rows.Add($"{sfz}-{sj}-{sj2}-{kx}-{yz}");
                    if (kx?.Contains("是") == true && yz?.Contains("否") == true)
                    {
                        actionCount++;
                    }
                    else
                    {
                        notMapCount++;
                    }
                    totalRowCount++;
                }
            }
            Console.WriteLine($"已读行数: {totalRowCount} , 符合条件的数量: {actionCount}/{notMapCount}");
        }
    }
}