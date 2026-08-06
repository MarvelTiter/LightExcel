using LightExcel.OpenXml;
using LightExcel.TypedDeserializer;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace LightExcel;

internal partial class ExcelReader
{
#if NET8_0_OR_GREATER
    IAsyncEnumerator<Row>? rowAsyncEnumerator;
    IAsyncEnumerator<Sheet>? sheetAsyncEnumerator;

    internal IAsyncEnumerator<Sheet> SheetAsyncEnumerator
    {
        get
        {
            return sheetAsyncEnumerator ??= document.WorkBook.WorkSheets.GetAsyncEnumerator();
        }
    }

    async Task<bool> HasRowsAsync()
    {
        // TODO startRow更改过，需要重新评估这里的代码
        if (startRow == 0)
        {
            if (rowAsyncEnumerator is null)
            {
                return false;
            }
            return await rowAsyncEnumerator.MoveNextAsync();
        }
        else
        {
            if (rowAsyncEnumerator is null)
                return false;

            while (await rowAsyncEnumerator.MoveNextAsync())
            {
                if (rowAsyncEnumerator.Current.RowIndex >= startRow)
                {
                    return true;
                }
            }
            return false;
        }
    }

    public async Task<bool> ReadAsync()
    {
        if (await HasRowsAsync())
        {
            //if (startColumn > 1)
            //    cells = rowEnumerator!.Current.RowDatas.Skip(startColumn - 1);
            //else
            if (startColumn > 1)
            {
                fixedColumn = startColumn - 1;
            }
            cells = rowAsyncEnumerator!.Current!.Children;
            System.Diagnostics.Debug.WriteLine($"{RowIndex}: Async Read, Cells {cells.Count}");
            return true;
        }
        return false;
    }

    async Task<bool> HasSheetAsync()
    {
        if (targetSheet == null)
        {
            return await SheetAsyncEnumerator.MoveNextAsync();
        }
        else
        {
            while (await SheetAsyncEnumerator.MoveNextAsync())
            {
                if (SheetAsyncEnumerator.Current.Name == targetSheet)
                {
                    return true;
                }
            }
            return false;
        }
    }

    public async Task<bool> NextResultAsync()
    {
        if (await HasSheetAsync())
        {
            rowAsyncEnumerator = SheetAsyncEnumerator.Current.GetAsyncEnumerator();
            if (configuration.UseHeader && await rowAsyncEnumerator.MoveNextAsync())
            {
                //Func<Cell, string> action = configuration.TrimHeader ? WithTrim : Original;
                while (rowAsyncEnumerator.Current.RowIndex < headerRowIndex)
                {
                    if (await rowAsyncEnumerator.MoveNextAsync())
                        continue;
                    else
                        break;
                }
                heads = [.. rowAsyncEnumerator.Current.Children.Select(Original)];
                if (startRow == 1)
                    startRow += headerRowIndex;
            }
            return true;
        }
        return false;
    }

    public async IAsyncEnumerable<dynamic> AsDynamicAsync([EnumeratorCancellation]CancellationToken cancellationToken = default)
    {
        Func<IExcelDataReader, object>? deserializer = null;
        while (await ReadAsync())
        {
            if (cancellationToken.IsCancellationRequested)
            {
                yield break;
            }
            deserializer ??= DynamicDeserialize.GetMapperRowDeserializer(this, startColumn);
            yield return deserializer.Invoke(this);
        }
    }

    public async IAsyncEnumerable<TData> AsTypedAsync<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.All)]
    TData>([EnumeratorCancellation]CancellationToken cancellationToken = default)
    {
        while (await ReadAsync())
        {
            if (cancellationToken.IsCancellationRequested)
            {
                yield break;
            }
            yield return ExpressionDeserialize<TData>.Deserialize(this);
        }
    }

#endif
}
