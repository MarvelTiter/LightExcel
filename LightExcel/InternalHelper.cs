using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

internal static class InternalHelper
{
    internal static Cell CreateTypedCell(Type type, object value)
    {
        var cell = new Cell();
        if (type == typeof(bool))
        {
            cell.CellValue = new CellValue((bool)value);
            cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
        }
        else if (type == typeof(DateTime))
        {
            cell.CellValue = new CellValue((DateTime)value);
            cell.DataType = new EnumValue<CellValues>(CellValues.Date);
        }
        else if (type == typeof(DateTimeOffset))
        {
            cell.CellValue = new CellValue((DateTimeOffset)value);
            cell.DataType = new EnumValue<CellValues>(CellValues.Date);
        }
        else if (value.IsNumeric<decimal>(out var v))
        {
            cell.CellValue = new CellValue((decimal)v);
            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (value.IsNumeric<double>(out var v1))
        {
            cell.CellValue = new CellValue((double)v1);
            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (value.IsNumeric<int>(out var v2))
        {
            cell.CellValue = new CellValue((int)v2);
            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else
        {
            cell.CellValue = new CellValue(value?.ToString() ?? "");
            cell.DataType = new EnumValue<CellValues>(CellValues.String);
        }

        return cell;
    }
    internal static bool IsNumeric<T>(this object? obj, out object value) where T : struct
    {
        value = null;
        if (obj == null)
        {
            return false;
        }
        if (typeof(T) == typeof(decimal))
        {
            if (decimal.TryParse(obj.ToString(), out var d))
            {
                value = d;
                return true;
            }
        }

        if (typeof(T) == typeof(double))
        {
            if (double.TryParse(obj.ToString(), out var d))
            {
                value = d;
                return true;
            }
        }

        if (typeof(T) == typeof(int))
        {
            if (int.TryParse(obj.ToString(), out var d))
            {
                value = d;
                return true;
            }
        }
        return false;
    }
}
