using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using LightExcel.OpenXml;
using LightExcel.TypedDeserializer;
using LightExcel.Utils;

namespace LightExcel
{
    public partial class ExcelHelper : IExcelHelper
    {
        private readonly ExcelConfiguration configuration = new ExcelConfiguration();

        public IExcelDataReader ReadExcel(string path, string? sheetName = null, Action<ExcelConfiguration>? config = null)
        {
            config?.Invoke(configuration);
            var archive = ExcelDocument.Open(path, configuration);
            return new ExcelReader(archive, configuration, sheetName);
        }

        public IEnumerable<T> QueryExcel<T>(string path, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
        {
            using var reader = ReadExcel(path, sheetName, config);
            while (reader.NextResult())
            {
                while (reader.Read())
                {
                    yield return ExpressionDeserialize<T>.Deserialize(reader);
                }
            }
        }

        public IEnumerable<dynamic> QueryExcel(string path, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
        {
            using var reader = ReadExcel(path, sheetName, config);
            while (reader.NextResult())
            {
                foreach (var item in reader.AsDynamic())
                {
                    yield return item;
                }
            }
        }

        public void WriteExcel(string path, object data, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
        {
            if (File.Exists(path)) File.Delete(path);
            config?.Invoke(configuration);
            using var trans = new TransExcelHelper(path, configuration);
            trans.WriteExcel(data, sheetName);
        }
        public void WriteExcelByTemplate(string path, string template, object data, string sheetName = "Sheet1", Action<ExcelConfiguration>? config = null)
        {
            config?.Invoke(configuration);
            HandleWriteTemplate(path, template, data, sheetName);
        }

        public ITransactionExcelHelper BeginTransaction(string path, Action<ExcelConfiguration>? config = null)
        {
            if (File.Exists(path)) File.Delete(path);
            config?.Invoke(configuration);
            return new TransExcelHelper(path, configuration);
        }

        
    }
}
