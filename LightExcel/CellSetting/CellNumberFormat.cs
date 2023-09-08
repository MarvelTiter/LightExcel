using LightExcel.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.CellSetting
{
    internal class CellNumberFormat : IExcelCellStyle
    {
        private readonly Func<object?, bool>? filter;
        public FormatType Type => FormatType.Number;
        public CellNumberFormat(NumberFormat format, Func<object?, bool>? filter = null)
        {
            Format = format;
            this.filter = filter;
        }

        public NumberFormat Format { get; }

        public bool HasStyle(object? value)
        {
            return filter?.Invoke(value) ?? true;
        }
    }
}
