using LightExcel.Extensions;
using LightExcel.OpenXml;
using LightExcel.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

        int startColumn = 0;
        int startRow = 0;
        public ExcelReader(ExcelArchiveEntry document, ExcelConfiguration configuration, string? targetSheet = null)
        {
            this.document = document;
            this.configuration = configuration;
            this.targetSheet = targetSheet;
            sheetEnumerator = document.WorkBook.WorkSheets.GetEnumerator();
            var (X, Y) = ReferenceHelper.ConvertCellReferenceToXY(configuration.StartCell);
            startColumn = X ?? 0;
            startRow = Y ?? 0;
        }

        private Cell? ElementAt(int i)
        {
            if (i > cells.Length) return null;
            return cells[i];
        }

        private Cell? CellAt(int i)
        {
            if (i > cells.Length) return null;
            return cells[i];
        }


        public bool GetBoolean(int i)
        {
            bool ret = false;
            ElementAt(i)?.TryGetBoolean(Sst, out ret);
            return ret;
        }

        public DateTime GetDateTime(int i)
        {
            DateTime ret = DateTime.MinValue;
            ElementAt(i)?.TryGetDateTime(Sst, out ret);
            return ret;
        }

        public decimal GetDecimal(int i)
        {
            decimal ret = default;
            ElementAt(i)?.TryGetDecimal(Sst, out ret);
            return ret;
        }

        public double GetDouble(int i)
        {
            double ret = default;
            ElementAt(i)?.TryGetDouble(Sst, out ret);
            return ret;
        }


        public int GetInt32(int i)
        {
            int ret = default;
            ElementAt(i)?.TryGetInt(Sst, out ret);
            return ret;
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
            return ElementAt(i)?.GetCellValue(document.WorkBook.SharedStrings) ?? "";
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
                if (configuration.UseHeader && configuration.StartCell == null && rowEnumerator.MoveNext())
                {
                    heads = rowEnumerator.Current.RowDatas.Select(c => c.GetCellValue(document.WorkBook.SharedStrings) ?? "").ToArray();
                }
                return true;
            }
            return false;
        }

        bool HasRows
        {
            get
            {
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
    }
}
