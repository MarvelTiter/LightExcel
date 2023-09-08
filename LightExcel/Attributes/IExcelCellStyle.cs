using DocumentFormat.OpenXml;
using LightExcel.CellSetting;

namespace LightExcel.Attributes
{
    //internal interface IExcelCellSetting<T>
    public abstract class ExcelCellSetting : Attribute
    {
        internal abstract IExcelCellStyle CreateElement();
        //bool ValueFilter(T value);
    }
}