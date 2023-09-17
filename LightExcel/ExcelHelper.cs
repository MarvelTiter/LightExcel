using System.Data;
using LightExcel.OpenXml;

namespace LightExcel
{
    public partial class ExcelHelper
    {
        void InternalWriteExcelWithTemplate(string path, string template, object data)
        {

        }
    }
    public partial class ExcelHelper : IExcelHelper
    {
        private readonly ExcelHelperConfiguration configuration = new ExcelHelperConfiguration();
        const string DEFAULT_SHEETNAME = "sheet";
        //public void WriteExcel(string path, object data, string? sheetName = "sheet", bool appendSheet = true)
        //{
        //    try
        //    {
        //        configuration.AllowAppendSheet = appendSheet;
        //        InternalWriteExcel(path, data, sheetName);
        //    }
        //    catch (Exception)
        //    {
        //        throw;
        //    }
        //}

        public void WriteExcel(string path, string template, object data)
        {
            try
            {
                InternalWriteExcelWithTemplate(path, template, data);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public IExcelDataReader ReadExcel(string path)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<T> QueryExcel<T>(string path, string? sheetName = null, int startRow = 2)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<dynamic> QueryExcel(string path, string? sheetName = null, int startRow = 2)
        {
            throw new NotImplementedException();
        }

        private void InternalWriteExcel(string path, object data, string? sheetName)
        {
            using var doc = GetDocument(path);
            var dataType = data.GetType();
            var render = RenderProvider.GetDataRender(dataType);

            WriteSheet(doc, data, render, sheetName);

            doc.Save();
        }

        private ExcelArchiveEntry GetDocument(string path)
        {
            ExcelArchiveEntry? doc = null;
            try
            {
                if (File.Exists(path))
                {
                    // 文件存在并且，允许追加Sheet
                    doc = ExcelDocument.Open(path, configuration);

                }
                else
                {
                    File.Delete(path);
                    doc = ExcelDocument.Create(path, configuration);
                }
            }
            catch (Exception)
            {
                throw;
            }
            return doc;
        }

        private void WriteSheet(ExcelArchiveEntry doc, object data, IDataRender render, string? sheetName)
        {
            sheetName = DEFAULT_SHEETNAME ?? string.Empty;
            var sheet = doc.WorkBook!.AddNewSheet(sheetName);

            //创建表头
            CreateHeader(sheet, data, render);

            //创建内容数据
            CreateBody(sheet, data, render);

        }

        /// <summary>
        /// 创建表头
        /// </summary>
        /// <param name="worksheetPart">WorksheetPart 对象</param>
        private void CreateHeader(Sheet sheet, object data, IDataRender render)
        {
            //var heads = render.RenderHeader(data);
            //sheet.AppendChild(heads);
        }

        private void CreateBody(Sheet sheet, object data, IDataRender render, int rowIndex = 2)
        {
            //var rows = render.RenderBody(data);
            int startIndex = rowIndex;
            //foreach (var r in rows)
            //{
            //    r.RowIndex = startIndex;
            //    //sheet.AppendChild(r);
            //    startIndex++;
            //}
        }

        public void WriteExcel(string path, object data, string sheetName = "sheet", Action<ExcelHelperConfiguration>? action = null)
        {
            throw new NotImplementedException();
        }

        public ITransactionExcelHelper BeginTransaction(string path, Action<ExcelHelperConfiguration>? config = null)
        {
            config?.Invoke(configuration);
            return new TransExcelHelper(path, configuration);
        }
    }
}
