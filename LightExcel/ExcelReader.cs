using LightExcel.OpenXml;
using LightExcel.TypedDeserializer;
using LightExcel.Utils;
using System.Data;
using System.Reflection.PortableExecutable;

namespace LightExcel
{
    internal class ExcelReader : IExcelDataReader
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

        public int FieldCount => cells.Length;

        Cell[] cells = Array.Empty<Cell>();
        string[] heads = Array.Empty<string>();

        int startColumn = 1;
        int startRow = 1;
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
            if (i < 0 || i > cells.Length) return null;
            return cells[i];
        }


        public bool GetBoolean(int i)
        {
            bool ret = false;
            CellAt(i)?.TryGetBoolean(Sst, out ret);
            return ret;
        }

        public DateTime GetDateTime(int i)
        {
            DateTime ret = DateTime.MinValue;
            CellAt(i)?.TryGetDateTime(Sst, out ret);
            return ret;
        }

        public decimal GetDecimal(int i)
        {
            decimal ret = default;
            CellAt(i)?.TryGetDecimal(Sst, out ret);
            return ret;
        }

        public double GetDouble(int i)
        {
            double ret = default;
            CellAt(i)?.TryGetDouble(Sst, out ret);
            return ret;
        }


        public int GetInt32(int i)
        {
            int ret = default;
            CellAt(i)?.TryGetInt(Sst, out ret);
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
                    heads = rowEnumerator.Current.RowDatas.Select(c => c.GetCellValue(Sst) ?? "").ToArray();
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
                if (startColumn > 1)
                    cells = rowEnumerator!.Current.RowDatas.Skip(startColumn - 1).ToArray();
                else
                    cells = rowEnumerator!.Current.RowDatas.ToArray();
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
}
