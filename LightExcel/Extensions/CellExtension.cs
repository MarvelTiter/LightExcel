using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.Extensions
{
    internal static class CellExtension
    {
        /// <summary>
        /// 获取单位格的值
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="workbookPart"></param>
        /// <param name="type">1 不去空格 2 前后空格 3 所有空格  </param>
        /// <returns></returns>
        public static string? GetCellValue(this Cell cell, WorkbookPart? workbookPart)
        {
            //合并单元格不做处理
            if (cell.CellValue == null)
                return string.Empty;

            string cellInnerText = cell.CellValue.InnerXml;

            //纯字符串
            if (cell.DataType != null && (cell.DataType.Value == CellValues.SharedString || cell.DataType.Value == CellValues.String || cell.DataType.Value == CellValues.Number))
            {
                //获取spreadsheetDocument中共享的数据
                SharedStringTable? stringTable = workbookPart?.SharedStringTablePart?.SharedStringTable;

                //如果共享字符串表丢失，则说明出了问题。
                if (!stringTable?.Any() ?? true)
                    return string.Empty;

                string? text = stringTable?.ElementAt(int.Parse(cellInnerText)).InnerText;
                return text;
            }
            //bool类型
            else if (cell.DataType != null && cell.DataType.Value == CellValues.Boolean)
            {
                return (cellInnerText != "0").ToString().ToUpper();
            }
            //数字格式代码（numFmtId）小于164是内置的：https://www.it1352.com/736329.html
            else
            {
                //为空为数值
                if (!cell.StyleIndex?.HasValue ?? true)
                    return cellInnerText;

                Stylesheet? styleSheet = workbookPart?.WorkbookStylesPart?.Stylesheet;
                CellFormat? cellFormat = styleSheet?.CellFormats?.ChildElements[(int)cell.StyleIndex!.Value] as CellFormat;

                uint? formatId = cellFormat?.NumberFormatId?.Value;
                double doubleTime;//OLE 自动化日期值
                DateTime dateTime;//yyyy/MM/dd HH:mm:ss
                switch (formatId)
                {
                    case 0://常规
                        return cellInnerText;
                    case 9://百分比【0%】
                    case 10://百分比【0.00%】
                    case 11://科学计数【1.00E+02】
                    case 12://分数【1/2】
                        return cellInnerText;
                    case 14:
                        doubleTime = double.Parse(cellInnerText);
                        dateTime = DateTime.FromOADate(doubleTime);
                        return dateTime.ToString("yyyy/MM/dd");
                    //case 15:
                    //case 16:
                    case 17:
                        doubleTime = double.Parse(cellInnerText);
                        dateTime = DateTime.FromOADate(doubleTime);
                        return dateTime.ToString("yyyy/MM");
                    //case 18:
                    //case 19:
                    case 20:
                        doubleTime = double.Parse(cellInnerText);
                        dateTime = DateTime.FromOADate(doubleTime);
                        return dateTime.ToString("H:mm");
                    case 21:
                        doubleTime = double.Parse(cellInnerText);
                        dateTime = DateTime.FromOADate(doubleTime);
                        return dateTime.ToString("HH:mm:ss");
                    case 22:
                        doubleTime = double.Parse(cellInnerText);
                        dateTime = DateTime.FromOADate(doubleTime);
                        return dateTime.ToString("yyyy/MM/dd HH:mm");
                    //case 45:
                    //case 46:
                    case 47:
                        doubleTime = double.Parse(cellInnerText);
                        dateTime = DateTime.FromOADate(doubleTime);
                        return dateTime.ToString("yyyy/MM/dd");
                    case 58://【中国】11月11日
                        doubleTime = double.Parse(cellInnerText);
                        dateTime = DateTime.FromOADate(doubleTime);
                        return dateTime.ToString("MM/dd");
                    case 176://【中国】2020年11月11日
                        doubleTime = double.Parse(cellInnerText);
                        dateTime = DateTime.FromOADate(doubleTime);
                        return dateTime.ToString("yyyy/MM/dd");
                    case 177://【中国】11:22:00
                        doubleTime = double.Parse(cellInnerText);
                        dateTime = DateTime.FromOADate(doubleTime);
                        return dateTime.ToString("HH:mm:ss");
                    default:
                        return cellInnerText;
                }
            }
        }
    }
}
