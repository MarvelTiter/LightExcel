using LightExcel.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.CellSetting
{
    internal class CellFillStyle : IExcelCellStyle
    {
        private readonly Func<object?, bool>? filter;
        public FormatType Type => FormatType.Fill;
        public FillType FillType { get; set; } = FillType.Solid;
        public string? BackgroundColor { get; set; }
        public string? ForegroundColor { get; set; }

        public CellFillStyle(FillType fillType, string? bg, string? fg, Func<object?,bool>? filter = null)
        {
            FillType = fillType;
            BackgroundColor = bg;
            ForegroundColor = fg;
            this.filter = filter;
        }

        public bool HasStyle(object? value)
        {
            return filter?.Invoke(value) ?? true;
        }
    }
}
