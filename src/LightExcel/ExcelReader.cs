using LightExcel.OpenXml;
using LightExcel.TypedDeserializer;
using LightExcel.Utils;
using System.Diagnostics.CodeAnalysis;

namespace LightExcel
{

    internal partial class ExcelReader : IExcelDataReader
    {
        private bool disposedValue;
        private readonly ExcelArchiveEntry document;
        private readonly ExcelConfiguration configuration;
        private readonly string? targetSheet;

        //private readonly IEnumerable<WorksheetPart> sheetParts
        SharedStringTable? Sst => document.WorkBook.SharedStrings;
        public string? this[int i] => CellAt(i)?.GetCellValue(Sst);
        public string? this[string name] => CellAt(GetOrdinal(name))?.GetCellValue(Sst);

        IEnumerator<Row>? rowEnumerator;
        IEnumerator<Sheet> sheetEnumerator;

        public string CurrentSheetName => sheetEnumerator.Current?.Name?.ToString() ?? "";

        public int FieldCount => cells.Count;
        public int RowIndex => rowEnumerator?.Current.RowIndex ?? 0;
        List<Cell> cells = [];
        string[] heads = [];

        int startColumn = 1;
        int startRow = 1;
        int fixedColumn = 0;
        public ExcelReader(ExcelArchiveEntry document, ExcelConfiguration configuration, string? targetSheet = null)
        {
            this.document = document;
            this.configuration = configuration;
            this.targetSheet = targetSheet;
            sheetEnumerator = document.WorkBook.WorkSheets.GetEnumerator();
            var (X, Y) = ReferenceHelper.ConvertCellReferenceToXY(configuration.StartCell);
            startColumn = X ?? 1;
            startRow = Y ?? 1;
        }


        private Cell? CellAt(int i)
        {
            //System.Diagnostics.Debug.WriteLine($"{RowIndex}: CellAt {i}, Enable {i + fixedColumn < 0 || i + fixedColumn >= cells.Count}");
            if (i + fixedColumn < 0 || i + fixedColumn >= cells.Count)
                return null;

            return cells[i + fixedColumn];
        }

        public bool GetBoolean(int i)
        {
            bool ret = false;
            CellAt(i)?.TryGetBoolean(Sst, i, configuration, out ret);
            return ret;
        }

        public DateTime GetDateTime(int i)
        {
            DateTime ret = DateTime.MinValue;
            CellAt(i)?.TryGetDateTime(Sst, i, configuration, out ret);
            return ret;
        }

        public decimal GetDecimal(int i)
        {
            decimal ret = default;
            CellAt(i)?.TryGetDecimal(Sst, i, configuration, out ret);
            return ret;
        }

        public double GetDouble(int i)
        {
            double ret = default;
            CellAt(i)?.TryGetDouble(Sst, i, configuration, out ret);
            return ret;
        }


        public int GetInt32(int i)
        {
            int ret = default;
            CellAt(i)?.TryGetInt(Sst, i, configuration, out ret);
            return ret;
        }

        public bool IsNullOrEmpty(int i)
        {
            return string.IsNullOrEmpty(CellAt(i)?.Value);
        }


        public string GetName(int i)
        {
            if (i < heads.Length)
            {
                return heads[i];
            }
            return "";
        }

        public int GetOrdinal(string name)
        {
            return Array.IndexOf(heads, name);
        }

        public string GetValue(int i)
        {
            return CellAt(i)?.GetCellValue(Sst) ?? "";
        }

        bool HasSheet
        {
            get
            {
                if (targetSheet == null)
                    return sheetEnumerator.MoveNext();
                else
                {
                    while (sheetEnumerator.MoveNext())
                    {
                        if (sheetEnumerator.Current.Name == targetSheet)
                        {
                            return true;
                        }
                    }
                    return false;
                }
            }
        }

        public bool NextResult()
        {
            if (HasSheet)
            {
                rowEnumerator = sheetEnumerator.Current.GetEnumerator();
                if (configuration.UseHeader && rowEnumerator.MoveNext())
                {
                    heads = [.. rowEnumerator.Current.Children.Select(c => c.GetCellValue(Sst) ?? "")];
                    startRow += 1;
                }
                return true;
            }
            return false;
        }

        bool HasRows
        {
            get
            {
                // TODO startRow更改过，需要重新评估这里的代码
                if (startRow == 0)
                    return rowEnumerator?.MoveNext() ?? false;
                else
                {
                    while (rowEnumerator?.MoveNext() ?? false)
                    {
                        if (rowEnumerator.Current.RowIndex >= startRow)
                        {
                            return true;
                        }
                    }
                    return false;
                }
            }
        }

        public bool Read()
        {
            if (HasRows)
            {
                //if (startColumn > 1)
                //    cells = rowEnumerator!.Current.RowDatas.Skip(startColumn - 1);
                //else
                if (startColumn > 1)
                {
                    fixedColumn = startColumn - 1;
                }
                cells = rowEnumerator!.Current!.Children;
                //System.Diagnostics.Debug.WriteLine($"{RowIndex}: Read, Cells {cells.Count}");
                return true;
            }
            return false;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    document?.Dispose();
                }
                disposedValue = true;
            }
        }
        public void Close()
        {

        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
        /// <summary>
        /// <para>根据Excel列取值</para>
        /// <para>eg: </para>
        /// <code>var datas = reader.AsDynamic();</code>
        /// <code>var a = datas.A; </code>
        /// <code>var b = datas.B; </code>
        /// </summary>
        /// <returns></returns>
        public IEnumerable<dynamic> AsDynamic()
        {
            Func<IExcelDataReader, object>? deserializer = null;
            while (Read())
            {
                deserializer ??= DynamicDeserialize.GetMapperRowDeserializer(this, startColumn);
                yield return deserializer.Invoke(this);
            }
        }

        public IEnumerable<TData> AsTyped<TData>()
        {
            while (Read())
            {
                yield return ExpressionDeserialize<TData>.Deserialize(this);
            }
        }
    }

    internal partial class ExcelReader
    {
        //object System.Data.IDataRecord.this[int i] => throw new NotImplementedException();

        //object System.Data.IDataRecord.this[string name] => throw new NotImplementedException();
        //object System.Data.IDataRecord.GetValue(int i)
        //{
        //    return GetValue(i);
        //}

        public int Depth => throw new NotImplementedException();

        public bool IsClosed => throw new NotImplementedException();

        public int RecordsAffected => throw new NotImplementedException();


        public byte GetByte(int i)
        {
            throw new NotImplementedException();
        }

        public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length)
        {
            throw new NotImplementedException();
        }

        public char GetChar(int i)
        {
            throw new NotImplementedException();
        }

        public long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length)
        {
            throw new NotImplementedException();
        }

        public System.Data.IDataReader GetData(int i)
        {
            throw new NotImplementedException();
        }

        public string GetDataTypeName(int i)
        {
            throw new NotImplementedException();
        }

        public Type GetFieldType(int i)
        {
            throw new NotImplementedException();
        }

        public float GetFloat(int i)
        {
            throw new NotImplementedException();
        }

        public Guid GetGuid(int i)
        {
            throw new NotImplementedException();
        }

        public short GetInt16(int i)
        {
            throw new NotImplementedException();
        }

        public long GetInt64(int i)
        {
            throw new NotImplementedException();
        }

        public System.Data.DataTable? GetSchemaTable()
        {
            throw new NotImplementedException();
        }

        public string GetString(int i)
        {
            throw new NotImplementedException();
        }

        public int GetValues(object[] values)
        {
            throw new NotImplementedException();
        }

        public bool IsDBNull(int i)
        {
            throw new NotImplementedException();
        }


    }
}
