﻿using LightExcel.OpenXml;
using LightExcel.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace LightExcel
{
    public partial class ExcelHelper
    {
        /// <summary>
        /// 仅支持第一个sheet
        /// </summary>
        /// <param name="path"></param>
        /// <param name="template"></param>
        /// <param name="data"></param>
        /// <param name="sheetName"></param>
        /// <exception cref="Exception"></exception>
        internal void HandleWriteTemplate(ExcelArchiveEntry doc, object data, string sheetName)
        {
            configuration.FillByTemplate = true;
            
            // 获取sheet对象
            var sheet = doc.WorkBook.WorkSheets.FirstOrDefault() ?? throw new Exception("read excel sheet failed");
            // 获取最后一行当模板
            var templateRow = sheet.ToList().Last();
            // 获取共享字符串列表
            var sst = doc.WorkBook.SharedStrings?.ToList();
            var render = RenderProvider.GetDataRender(data.GetType(), configuration);
            var columns = configuration.FillWithPlacholder ? CollectExcelColumnInfos(templateRow, sst).ToArray() : render.CollectExcelColumnInfo(data).ToArray();
            sheet.Columns = columns;
            if (configuration.FillWithPlacholder)
            {
                templateRow.IsTemplateRow = true;
                configuration.StartRowIndex = templateRow.RowIndex - 1;
            }
            else
            {
                configuration.StartRowIndex = templateRow.RowIndex;
            }
            var newRows = render.RenderBody(data, sheet, columns, new TransConfiguration { SheetNumberFormat = configuration.AddSheetNumberFormat });
            sheet.Replace(sheet.Concat(newRows));
            doc.Save();
        }

        static readonly Regex extract = new Regex("{{(.+)}}");
        private IEnumerable<ExcelColumnInfo> CollectExcelColumnInfos(Row templateRow, List<SharedStringNode>? sst)
        {
            foreach (var cell in templateRow.RowDatas)
            {
                string? name = cell.Value;
                var (X, Y) = ReferenceHelper.ConvertCellReferenceToXY(cell.Reference);
                if (cell.Type == "s")
                {
                    if (int.TryParse(name, out var s) && sst!.Count > s)
                    {
                        name = sst[s].Content;
                    }
                }
                if (name != null)
                {
                    var match = extract.Match(name);
                    if (match.Success)
                    {
                        name = match.Groups[1].Value;
                        var col = new ExcelColumnInfo(name) { ColumnIndex = X ?? 0, StyleIndex = cell.StyleIndex };
                        yield return col;
                    }
                }
            }
        }
    }
}
