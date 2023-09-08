using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LightExcel.CellSetting;
using LightExcel.Enums;

namespace LightExcel
{
    internal partial class StyleHelper
    {
        static void CheckAndCreateStylePart(WorkbookPart workbookPart)
        {
            if (workbookPart.WorkbookStylesPart == null)
            {
                var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                var styleSheet = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
                styleSheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                styleSheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
                stylesPart.Stylesheet = styleSheet;
            }
        }
        internal static uint? GetExcelNumberFormatId(WorkbookPart workbookPart, CellNumberFormat? cnf)
        {
            if (cnf == null) return null;
            CheckAndCreateStylePart(workbookPart);
            Stylesheet stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet;
            var code = cnf.Format.GetFormatCode();
            if (stylesheet.NumberingFormats == null)
            {
                stylesheet.NumberingFormats = new();
                var format = new NumberingFormat
                {
                    NumberFormatId = UInt32Value.FromUInt32(0),
                    FormatCode = code
                };
                stylesheet.NumberingFormats.AppendChild(format);
                stylesheet.NumberingFormats.Count = 1;
                return 0;
            }
            else
            {
                foreach (NumberingFormat item in stylesheet.NumberingFormats.Cast<NumberingFormat>())
                {
                    if (item.FormatCode == code)
                        return item.NumberFormatId!.Value;
                }
                var count = stylesheet.NumberingFormats.Count!.Value;
                var format = new NumberingFormat
                {
                    NumberFormatId = UInt32Value.FromUInt32(count),
                    FormatCode = code
                };
                stylesheet.NumberingFormats.AppendChild(format);
                stylesheet.NumberingFormats.Count = count + 1;
                return count;
            }
        }

        internal static uint? GetExcelFillId(WorkbookPart workbookPart, CellFillStyle? fill)
        {
            if (fill == null) return null;
            CheckAndCreateStylePart(workbookPart);
            Stylesheet stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet;
            // Fills的fillid=0和fillId=1的位置均是系统默认的。fillId=0的填充色是None，fillId=1的填充色是Gray125，但需要自定义填充色时，必须从fillId=2开始定义
            if (stylesheet.Fills == null)
            {
                stylesheet.Fills = new();
                //fillid=0
                Fill fillDefault = new(new PatternFill() { PatternType = PatternValues.None });
                stylesheet.Fills.AppendChild(fillDefault);

                //fillid=1
                Fill fillGray = new(new PatternFill() { PatternType = PatternValues.Gray125 });
                stylesheet.Fills.AppendChild(fillGray);
            }

            var index = 0u;
            foreach (Fill item in stylesheet.Fills.Cast<Fill>())
            {
                if (item.PatternFill!.PatternType! == (PatternValues)(int)fill.FillType
                    && item.PatternFill!.BackgroundColor?.Rgb == fill.BackgroundColor
                    && item.PatternFill!.ForegroundColor?.Rgb == fill.ForegroundColor)
                {
                    return index;
                }
                index++;
            }
            var fillElement = new Fill();
            var pattern = new PatternFill()
            {
                PatternType = (PatternValues)(int)fill.FillType,
            };
            if (fill.BackgroundColor != null)
            {
                pattern.BackgroundColor = new BackgroundColor { Rgb = fill.BackgroundColor };
            }
            if (fill.ForegroundColor != null)
            {
                pattern.ForegroundColor = new ForegroundColor { Rgb = fill.ForegroundColor };
            }
            fillElement.Append(pattern);
            stylesheet.Fills.AppendChild(fillElement);
            stylesheet.Fills.Count = index;
            return index;
        }

        internal static uint? GetExcelFormatId(WorkbookPart workbookPart, CellNumberFormat? num, CellFillStyle? fill)
        {
            var numId = GetExcelNumberFormatId(workbookPart, num);
            var fillId = GetExcelFillId(workbookPart, fill);
            if (!numId.HasValue && !fillId.HasValue)
            {
                return null;
            }
            Stylesheet stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet;
            if (stylesheet.CellFormats == null)
            {
                stylesheet.CellFormats = new();
                var cellFormat = new CellFormat
                {
                    FormatId = UInt32Value.FromUInt32(0)
                };
                if (numId.HasValue)
                {
                    cellFormat.NumberFormatId = numId;
                    cellFormat.ApplyNumberFormat = true;
                }
                if (fillId.HasValue)
                {
                    cellFormat.FillId = fillId;
                    cellFormat.ApplyFill = true;
                }
                stylesheet.CellFormats.AppendChild(cellFormat);
                stylesheet.CellFormats.Count = 1;
                stylesheet.Save();
                return 0;
            }
            else
            {
                foreach (CellFormat item in stylesheet.CellFormats.Cast<CellFormat>())
                {
                    if (item.NumberFormatId?.Value == numId && item.FillId?.Value == fillId)
                        return item.FormatId!.Value;
                }
                var count = stylesheet.CellFormats.Count!.Value;
                var format = new CellFormat
                {
                    FormatId = UInt32Value.FromUInt32(count)
                };
                if (numId.HasValue)
                {
                    format.NumberFormatId = numId;
                    format.ApplyNumberFormat = true;
                }
                if (fillId.HasValue)
                {
                    format.FillId = fillId;
                    format.ApplyFill = true;
                }
                stylesheet.CellFormats.AppendChild(format);
                stylesheet.CellFormats.Count = count + 1;
                stylesheet.Save();
                return count;
            }

        }

    }
}
