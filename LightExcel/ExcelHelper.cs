using System.Data;
using LightExcel.OpenXml;

namespace LightExcel
{

    public class ExcelHelper : IExcelHelper
    {
        private readonly ExcelConfiguration configuration = new ExcelConfiguration();
        

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


        public void WriteExcel(string path, object data, string sheetName = "sheet1", Action<ExcelConfiguration>? config = null)
        {
            if (File.Exists(path)) File.Delete(path);
            config?.Invoke(configuration);
            using var trans = new TransExcelHelper(path, configuration);
            trans.WriteExcel(data, sheetName);
        }
        public void WriteExcelByTemplate(string path, string template, object data, string sheetName = "sheet1", Action<ExcelConfiguration>? config = null)
        {
            config?.Invoke(configuration);
            using var doc = ExcelDocument.CreateByTemplate(path, template, configuration);
            foreach (var item in doc.WorkBook!.WorkSheets.First())
            {

            }
            foreach (var item in doc.WorkBook.SharedStrings)
            {

            }
        }

        public ITransactionExcelHelper BeginTransaction(string path, Action<ExcelConfiguration>? config = null)
        {
            if (File.Exists(path)) File.Delete(path);
            config?.Invoke(configuration);
            return new TransExcelHelper(path, configuration);
        }

       
    }
}
