using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public class ExcelColumnAttribute : Attribute
    {
        public string? Name { get; set; }
        public bool Ignore { get; set; }
        public bool NumberFormat { get; set; }
        public int? ColumnIndex { get; set; }
        public string? Format { get; set; }
        public bool AutoWidth { get; set; }
        public float? Width { get; set; }
    }
}
