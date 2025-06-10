using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using LightExcel.OpenXml;
using LightExcel.TypedDeserializer;
using LightExcel.Utils;

namespace LightExcel
{
    internal partial class ExcelHelper : IExcelHelper
    {
        //public ExcelConfiguration configuration { get; set; } = new();
        #region 读取
        public IExcelDataReader ReadExcel(string path, string? sheetName = null, Action<ExcelConfiguration>? config = null)
        {
            ExcelConfiguration configuration = new ExcelConfiguration();
            config?.Invoke(configuration);
            var archive = ExcelDocument.Open(path, configuration);
            return new ExcelReader(archive, configuration, sheetName);
        }

        public IEnumerable<T> QueryExcel<T>(string path, string? sheetName = null, Action<ExcelConfiguration>? config = null)
        {
            IExcelDataReader? reader = null;
            try
            {
                if (config is null)
                {
                    config = c => c.AddDynamicColumns(typeof(T).CollectEntityInfo());
                    reader = ReadExcel(path, sheetName, config);
                }
                else
                {
                    reader = ReadExcel(path, sheetName, c =>
                    {
                        config.Invoke(c);
                        c.AddDynamicColumns(typeof(T).CollectEntityInfo());
                    });
                }
                while (reader.NextResult())
                {
                    while (reader.Read())
                    {
                        yield return ExpressionDeserialize<T>.Deserialize(reader);
                    }
                }
            }
            finally
            {
                reader?.Dispose();
            }

        }

        public IEnumerable<dynamic> QueryExcel(string path, string? sheetName = null, Action<ExcelConfiguration>? config = null)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            using var reader = ReadExcel(path, sheetName, config);
            while (reader.NextResult())
            {
                foreach (var item in reader.AsDynamic())
                {
                    yield return item;
                }
            }
        }
        #endregion

        #region 写入

        public void WriteExcel(IDataRender render, string path, object data, string sheetName = "Sheet1", ExcelConfiguration? config = null)
        {
            config ??= new();
            if (File.Exists(path)) File.Delete(path);
            using var trans = new TransExcelHelper(path, config);
            trans.WriteExcel(render, data, sheetName);
        }

        public ITransactionExcelHelper BeginTransaction(ExcelArchiveEntry doc, ExcelConfiguration? config = null)
        {
            config ??= new();
            return new TransExcelHelper(doc, config);
        }


        #endregion
    }
}
