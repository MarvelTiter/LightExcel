using LightExcel.OpenXml;
using LightExcel.TypedDeserializer;
using LightExcel.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
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

        public IEnumerable<T> QueryExcel<
#if NET8_0_OR_GREATER
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.All)]
#endif
        T>(Stream stream, string? sheetName = null, Action<ExcelConfiguration>? config = null)
        {
            using var reader = ReadExcel(stream, sheetName, config);
            while (reader.NextResult())
            {
                while (reader.Read())
                {
                    yield return ExpressionDeserialize<T>.Deserialize(reader);
                }
            }
        }

        public IEnumerable<dynamic> QueryExcel(Stream stream, string? sheetName = null, Action<ExcelConfiguration>? config = null)
        {
            using var reader = ReadExcel(stream, sheetName, config);
            while (reader.NextResult())
            {
                foreach (var item in reader.AsDynamic())
                {
                    yield return item;
                }
            }
        }

#if NET8_0_OR_GREATER
        public async IAsyncEnumerable<T> QueryExcelAsync<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.All)] T>(Stream stream, string? sheetName = null, Action<ExcelConfiguration>? config = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            IExcelDataReader? reader = null;
            try
            {
                if (config is null)
                {
                    config = c =>
                    {
                        c.AddDynamicColumns(typeof(T).CollectEntityInfo());
                        c.EnableAsync = true;
                    };
                    reader = ReadExcel(stream, sheetName, config);
                }
                else
                {
                    reader = ReadExcel(stream, sheetName, c =>
                    {
                        config.Invoke(c);
                        c.EnableAsync = true;
                        c.AddDynamicColumns(typeof(T).CollectEntityInfo());
                    });
                }
                while (await reader.NextResultAsync())
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        yield break;
                    }
                    while (await reader.ReadAsync())
                    {
                        if (cancellationToken.IsCancellationRequested)
                        {
                            yield break;
                        }
                        yield return ExpressionDeserialize<T>.Deserialize(reader);
                    }
                }
            }
            finally
            {
                reader?.Dispose();
            }
        }

        public async IAsyncEnumerable<dynamic> QueryExcelAsync(Stream stream, string? sheetName = null, Action<ExcelConfiguration>? config = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            using var reader = ReadExcel(stream, sheetName, config);
            while (reader.NextResult())
            {
                await foreach (var item in reader.AsDynamicAsync(cancellationToken))
                {
                    yield return item;
                }
            }
        }
#endif
    }
}
