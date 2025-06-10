namespace TestProject1
{
    public class Model
    {
        public int Index { get; set; }
        public string? Name { get; set; }
        public DateTime Birthday { get; set; }
        [LightExcel.Attributes.ExcelColumn(Format = "yyyyMMdd")]
        public DateTime Birthday2 { get; set; }
    }
}