using LightExcel.OpenXml;
using LightExcel.Renders;
using LightExcel.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace LightExcel
{
    internal partial class TransExcelHelper : ITransactionExcelHelper
    {
        private bool disposedValue;
        private ExcelArchiveEntry? excelArchive;
        public ExcelConfiguration Configuration { get; set; }

        public TransExcelHelper(string path, ExcelConfiguration configuration)
        {
            if (File.Exists(path))
            {
                excelArchive = ExcelDocument.Open(path, configuration);
            }
            else
            {
                excelArchive = ExcelDocument.Create(path, configuration);
            }

            Configuration = configuration;
        }

        public TransExcelHelper(Stream stream, ExcelConfiguration configuration)
        {
            excelArchive = ExcelDocument.Create(stream, configuration);
            Configuration = configuration;
        }

        public TransExcelHelper(ExcelArchiveEntry doc, ExcelConfiguration configuration)
        {
            excelArchive = doc;
            Configuration = configuration;
        }

        public void WriteExcel(IDataRender render, object data, string? sheetName = null, TransConfiguration? config = null)
        {
            //WriteExcel(render, data, sheetName, config);
            config ??= new TransConfiguration(Configuration);
            Sheet sheet = excelArchive!.WorkBook.AddNewSheet(sheetName);
            render.Render(data, sheet, config);
        }

        internal static (List<Row>, Sheet) HandleTemplateHeader(ExcelArchiveEntry doc, string sheetName, ExcelConfiguration configuration)
        {
            configuration.FillByTemplate = true;
            // 获取sheet对象
            var sheet = doc.WorkBook.WorkSheets.FirstOrDefault() ?? throw new Exception("read excel sheet failed");
            // 获取最后一行当模板
            var header = sheet.ToList();
            var templateRow = header.Last();
            // 获取共享字符串列表
            var sst = doc.WorkBook.SharedStrings?.ToList();
            //var render = RenderProvider.GetDataRender(data.GetType(), configuration);
            if (configuration.FillWithPlacholder)
            {
                sheet.Columns = [.. CollectExcelColumnInfos(templateRow, sst)];
            }
            if (configuration.FillWithPlacholder)
            {
                templateRow.IsTemplateRow = true;
                configuration.StartRowIndex = templateRow.RowIndex - 1;
            }
            else
            {
                configuration.StartRowIndex = templateRow.RowIndex;
            }
            return (header, sheet);
        }

        /// <summary>
        /// 仅支持第一个sheet
        /// </summary>
        /// <param name="path"></param>
        /// <param name="template"></param>
        /// <param name="data"></param>
        /// <param name="sheetName"></param>
        /// <exception cref="Exception"></exception>
        internal static void WriteByTemplate(IDataRender render, ExcelArchiveEntry doc, object data, string sheetName, ExcelConfiguration configuration)
        {
            var (header, sheet) = HandleTemplateHeader(doc, sheetName, configuration);
            //var newRows = render.RenderBody(data, sheet, new TransConfiguration(configuration) { SheetNumberFormat = configuration.AddSheetNumberFormat });
            //sheet.Replace(header.Concat(newRows));
            render.SetCustomHeaders(header);
            sheet.DeleteEntry();
            render.Render(data, sheet, new TransConfiguration(configuration) { SheetNumberFormat = configuration.AddSheetNumberFormat });
            doc.Save();
        }

#if NET6_0_OR_GREATER

        public async Task WriteExcelAsync<TRender>(object data, string? sheetName = null, Action<TransConfiguration>? config = null, CancellationToken cancellationToken = default)
            where TRender : IAsyncDataRender
        {
            var render = AsyncRenderCreator<TRender>.Create(Configuration);
            var cfg = new TransConfiguration(Configuration);
            config?.Invoke(cfg);
            var sheet = excelArchive!.WorkBook.AddNewSheet(sheetName);
            await render.RenderAsync(data, sheet, cfg, cancellationToken);
        }

        internal static async Task WriteByTemplateAsync<TRender>(ExcelArchiveEntry doc, object data, string sheetName, ExcelConfiguration configuration, CancellationToken cancellationToken = default)
            where TRender : IAsyncDataRender
        {
            var render = AsyncRenderCreator<TRender>.Create(configuration);
            var (header, sheet) = HandleTemplateHeader(doc, sheetName, configuration);
            render.SetCustomHeaders(header);
            sheet.DeleteEntry();
            await render.RenderAsync(data, sheet, new TransConfiguration(configuration) { SheetNumberFormat = configuration.AddSheetNumberFormat }, cancellationToken);
            doc.Save();
        }

#endif


        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    excelArchive?.Save();
                    excelArchive?.Dispose();
                    excelArchive = null;
                }
                disposedValue = true;
            }
        }
        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

#if NET8_0_OR_GREATER
        [GeneratedRegex("{{(.+)}}")]
        private static partial Regex ExtractColumn();
#else
        static readonly Regex extract = new("{{(.+)}}");
        private static Regex ExtractColumn() => extract;
#endif
        private static IEnumerable<ExcelColumnInfo> CollectExcelColumnInfos(Row templateRow, List<SharedStringNode>? sst)
        {
            foreach (var cell in templateRow.Children)
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
                    var match = ExtractColumn().Match(name);
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
