using LightExcel.Renders;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel;

public static partial class ExcelHeperExtensions
{
    #region DataTableRender - DataTable作为数据源

    /// <summary>
    /// DataTable数据源-保存到文件
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcel(this IExcelHelper helper, string path, DataTable datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcel<DataTableRender>(helper, path, datas, sheetName, config);
    }

    /// <summary>
    /// DataTable数据源-使用模板文件
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcelByTemplate(this IExcelHelper helper, string path, string template, DataTable datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcelByTemplate<DataTableRender>(helper, path, template, datas, sheetName, config);
    }

    /// <summary>
    /// DataTable数据源-保存到内存流
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcel(this IExcelHelper helper, Stream stream, DataTable datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcel<DataTableRender>(helper, stream, datas, sheetName, config);
    }

    /// <summary>
    /// DataTable数据源-使用模板流保存到内存流
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcelByTemplate(this IExcelHelper helper, Stream stream, Stream templateStream, DataTable datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcelByTemplate<DataTableRender>(helper, stream, templateStream, datas, sheetName, config);
    }

    /// <summary>
    /// DataTable数据源-使用模板流保存到文件
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcelByTemplate(this IExcelHelper helper, string path, Stream templateStream, DataTable datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcelByTemplate<DataTableRender>(helper, path, templateStream, datas, sheetName, config);
    }

    /// <summary>
    /// DataTable数据源-使用模板文件保存到内存流
    /// </summary>
    /// <param name="helper"></param>
    /// <param name="path"></param>
    /// <param name="datas"></param>
    /// <param name="sheetName"></param>
    /// <param name="config"></param>
    public static void WriteExcelByTemplate(this IExcelHelper helper, Stream stream, string template, DataTable datas, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
    {
        InternalWriteExcelByTemplate<DataTableRender>(helper, stream, template, datas, sheetName, config);
    }

    #endregion
}
