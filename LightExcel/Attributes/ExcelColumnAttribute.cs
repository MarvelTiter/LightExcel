using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LightExcel.Enums;
using LightExcel.CellSetting;

namespace LightExcel.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        public string? Name { get; set; }
        public bool Ignore { get; set; }
    }

    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelStyleCellFillAttribute : ExcelCellSetting
    {
        public FillType Type { get; set; } = FillType.Solid;
        public string? BackgroundColor { get; set; }
        public string? ForegroundColor { get; set; }

        internal override IExcelCellStyle CreateElement()
        {
            return new CellFillStyle(Type, BackgroundColor, ForegroundColor);
        }
    }

    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelNumberFormatAttribute : ExcelCellSetting
    {
        public NumberFormat Format { get; }

        public ExcelNumberFormatAttribute(NumberFormat format)
        {
            Format = format;
        }

        internal override IExcelCellStyle CreateElement()
        {
            return new CellNumberFormat(Format);
        }
    }
}
