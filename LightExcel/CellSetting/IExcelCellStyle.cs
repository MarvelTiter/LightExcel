using LightExcel.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.CellSetting
{
    internal interface IExcelCellStyle
    {
        FormatType Type { get; }
        bool HasStyle(object? value);
    }
}
