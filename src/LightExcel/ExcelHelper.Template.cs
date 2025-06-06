using LightExcel.OpenXml;
using LightExcel.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace LightExcel
{
    internal partial class ExcelHelper
    {
        public void WriteExcelByTemplate(IDataRender render, ExcelArchiveEntry doc, object data, string sheetName = "Sheet1", ExcelConfiguration? config = null)
        {
            config ??= new ExcelConfiguration();
           TransExcelHelper.WriteByTemplate(render, doc, data, sheetName, config);
        }
        public void WriteExcelByTemplate(IDataRender render, string path, string template, object data, string sheetName = "Sheet1", ExcelConfiguration? config = null)
        {
            config ??= new();
            using var doc = ExcelDocument.CreateByTemplate(path, template, config);
            TransExcelHelper.WriteByTemplate(render, doc, data, sheetName, config);
        }

        public void WriteExcelByTemplate(IDataRender render, Stream stream, Stream templateStream, object data, string sheetName = "Sheet1", ExcelConfiguration? config = null)
        {
            config ??= new();
            using var doc = ExcelDocument.CreateByTemplate(stream, templateStream, config);
            TransExcelHelper.WriteByTemplate(render, doc, data, sheetName, config);
        }

    }
}
