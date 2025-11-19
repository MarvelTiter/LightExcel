using LightExcel.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
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
#if NET8_0_OR_GREATER
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
#endif
        public Type? Type { get; set; }
        public bool NumberFormat { get; set; }
        internal int ColumnIndex { get; set; }
        public string? StyleIndex { get; set; }
        public bool AutoWidth { get; set; }
        public double? Width { get; set; }
        internal ExcelColumnInfo(string name)
        {
            Name = name;
        }
    }
}
