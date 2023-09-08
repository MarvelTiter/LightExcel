using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using LightExcel.CellSetting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel
{
    internal partial class StyleHelper
    {
        List<IExcelCellStyle> excelCellStyles = new();

        internal void AddStyle(IExcelCellStyle style)
        {
            if (!excelCellStyles.Any(e => e.Type == style.Type))
            {
                excelCellStyles.Add(style);
            }
            var i = excelCellStyles.FindIndex(e => e.Type == style.Type);
            excelCellStyles.RemoveAt(i);
            excelCellStyles.Add(style);
        }

        internal bool HasStyle(object? value)
        {
            return excelCellStyles.Any(e => e.HasStyle(value));
        }

        internal uint? GetStyleIndex(WorkbookPart workbookPart)
        {
            var n = excelCellStyles.FirstOrDefault(e => e.Type == Enums.FormatType.Number) as CellNumberFormat;
            var f = excelCellStyles.FirstOrDefault(e => e.Type == Enums.FormatType.Fill) as CellFillStyle;
            return GetExcelFormatId(workbookPart, n, f);
        }
    }
}
