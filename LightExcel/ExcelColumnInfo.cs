using LightExcel.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel
{
    internal class ExcelColumnInfo
    {
        public string Name { get; }
        public bool Ignore { get; set; }
        public Property? Property { get; set; }
        public string? Format { get; set; }
        public ExcelColumnInfo(string name)
        {
            Name = name;
        }
    }
}
