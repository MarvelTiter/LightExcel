using System.Data;

namespace LightExcel
{
    public interface IExcelDataReader : IDataRecord, IDisposable
    {
        void Close();
        bool Read();
        bool NextResult();
    }
}
