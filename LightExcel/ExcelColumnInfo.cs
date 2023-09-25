using LightExcel.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel
{
    public sealed class ExcelColumnInfo
    {
        public string Name { get; }
        public bool Ignore { get; set; }
        public Property? Property { get; set; }
        public string? Format { get; set; }
        public Type? Type { get; set; }
        public bool NumberFormat { get; set; }
        internal int ColumnIndex {  get; set; }
        public bool AutoWidth { get; set; }
        internal ExcelColumnInfo(string name)
        {
            Name = name;
        }
    }
}
