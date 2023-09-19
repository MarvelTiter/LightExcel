using System.Data;
using System.Diagnostics.CodeAnalysis;

namespace LightExcel
{
    public interface IExcelDataReader : IDisposable
    {
        bool Read();
        bool NextResult();

        string? this[string name] { get; }
        string? this[int i] { get; }
        string CurrentSheetName { get; }
        int FieldCount { get; }
        bool GetBoolean(int i);
      
        DateTime GetDateTime(int i);
      
        decimal GetDecimal(int i);
       
        double GetDouble(int i);
      
        int GetInt32(int i);
               
        string GetName(int i);
      
        int GetOrdinal(string name);
       
        string GetValue(int i);
        bool IsNullOrEmpty(int i);
    }
}
