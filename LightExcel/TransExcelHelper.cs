using LightExcel.OpenXml;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel
{
    internal class TransExcelHelper : ITransactionExcelHelper
    {
        private bool disposedValue;
        private ExcelArchiveEntry? excelArchive;
        private readonly ExcelHelperConfiguration configuration;

        public TransExcelHelper(string path, ExcelHelperConfiguration configuration)
        {
            if (File.Exists(path))
            {
                excelArchive = ExcelDocument.Open(path, configuration);
            }
            else
            {
                excelArchive = ExcelDocument.Create(path, configuration);
            }

            this.configuration = configuration;
        }
        public void WriteExcel(object data, string? sheetName = null)
        {
            var sheet = excelArchive!.WorkBook.AddNewSheet(sheetName);
            var render = RenderProvider.GetDataRender(data.GetType());
            render.CollectExcelColumnInfo(data, configuration);
            var all = NeedToReaderRows(render, sheet, data);
            sheet!.Write(all);
        }

        private IEnumerable<Row> NeedToReaderRows(IDataRender render, Sheet sheet, object data)
        {
            if (configuration.UseHeader)
            {
                var header = render.RenderHeader(configuration);
                yield return header;
            }
            var datas = render.RenderBody(data, sheet, configuration);
            foreach (var row in datas)
            {
                yield return row;
            }
        }

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


    }
}
