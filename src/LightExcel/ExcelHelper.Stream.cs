using LightExcel.OpenXml;
using LightExcel.TypedDeserializer;
using LightExcel.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel
{
    internal partial class ExcelHelper
    {
        public void WriteExcel(IDataRender render, Stream stream, object data, string sheetName = "Sheet1", ExcelConfiguration? config = null)
        {
            config ??= new();
            using var trans = new TransExcelHelper(stream, config);
            trans.WriteExcel(render, data, sheetName);
        }
        
        public ITransactionExcelHelper BeginTransaction(Stream stream, Action<ExcelConfiguration>? config = null)
        {
            ExcelConfiguration configuration = new();
            config?.Invoke(configuration);
            return new TransExcelHelper(stream, configuration);
        }

        public IExcelDataReader ReadExcel(Stream stream, string? sheetName = null, Action<ExcelConfiguration>? config = null)
        {
            ExcelConfiguration configuration = new();
            config?.Invoke(configuration);
            var archive = ExcelDocument.Open(stream, configuration);
            return new ExcelReader(archive, configuration, sheetName);
        }

        public IEnumerable<T> QueryExcel<T>(Stream stream, string? sheetName, Action<ExcelConfiguration>? config = null)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            using var reader = ReadExcel(stream, sheetName, config);
            while (reader.NextResult())
            {
                while (reader.Read())
                {
                    yield return ExpressionDeserialize<T>.Deserialize(reader);
                }
            }
        }

        public IEnumerable<dynamic> QueryExcel(Stream stream, string? sheetName, Action<ExcelConfiguration>? config = null)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            using var reader = ReadExcel(stream, sheetName, config);
            while (reader.NextResult())
            {
                foreach (var item in reader.AsDynamic())
                {
                    yield return item;
                }
            }
        }
    }
}
