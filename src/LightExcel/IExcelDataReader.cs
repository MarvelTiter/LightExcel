using System.Data;
using System.Diagnostics.CodeAnalysis;

namespace LightExcel
{
    public interface IExcelDataReader : IDisposable
    {
        bool Read();
        bool NextResult();
#if NET8_0_OR_GREATER
        Task<bool> ReadAsync();
        Task<bool> NextResultAsync();
#endif
        string? this[string name] { get; }
        string? this[int i] { get; }
        int FieldCount { get; }

        bool GetBoolean(int i);
        DateTime GetDateTime(int i);
        decimal GetDecimal(int i);
        double GetDouble(int i);
        short GetInt16(int i);
        int GetInt32(int i);
        long GetInt64(int i);
        string GetName(int i);
        int GetOrdinal(string name);
        string GetValue(int i);

        bool IsNullOrEmpty(int i);
        string CurrentSheetName { get; }
        int RowIndex { get; }
        IEnumerable<dynamic> AsDynamic();
        IEnumerable<TData> AsTyped<
#if NET8_0_OR_GREATER
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.All)]
#endif
        TData>();

#if NET8_0_OR_GREATER
        IAsyncEnumerable<dynamic> AsDynamicAsync(CancellationToken cancellationToken = default);
        IAsyncEnumerable<TData> AsTypedAsync<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.All)]
        TData>(CancellationToken cancellationToken = default);
#endif
    }
}
