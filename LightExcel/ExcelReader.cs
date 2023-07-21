using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
        private readonly IEnumerable<WorksheetPart> sheetParts;
        private readonly IEnumerable<Sheet> sheets;
        private readonly SpreadsheetDocument document;
        public string? this[int i] => ElementAt(i)?.Text;
        public string? this[string name] => throw new NotImplementedException();

        SheetData? currentSheet;
        Row[] rows = Array.Empty<Row>();

        public string CurrentSheetName => sheets.ElementAt(currentSheetIndex)?.Name?.ToString() ?? "";

        public int FieldCount => cells.Length;

        Cell[] cells = Array.Empty<Cell>();
        string[] heads = Array.Empty<string>();
        int currentSheetIndex = -1;
        int currentRowIndex = -1;
        int headerRowIndex = 0;
        public ExcelReader(SpreadsheetDocument document, int start)
        {
            this.document = document;
            sheetParts = document.WorkbookPart?.WorksheetParts ?? Enumerable.Empty<WorksheetPart>();
            sheets = document.WorkbookPart?.Workbook.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>();
            currentRowIndex += start;
            headerRowIndex = start - 2;
        }

        private CellValue? ElementAt(int i)
        {
            if (i > cells.Length) return null;
            return cells[i].CellValue;
        }

        public void Close()
        {
            document?.Dispose();
        }

        public bool GetBoolean(int i)
        {
            bool ret = false;
            ElementAt(i)?.TryGetBoolean(out ret);
            return ret;
        }

        public DateTime GetDateTime(int i)
        {
            DateTime ret = DateTime.MinValue;
            ElementAt(i)?.TryGetDateTime(out ret);
            return ret;
        }

        public decimal GetDecimal(int i)
        {
            decimal ret = default;
            ElementAt(i)?.TryGetDecimal(out ret);
            return ret;
        }

        public double GetDouble(int i)
        {
            double ret = default;
            ElementAt(i)?.TryGetDouble(out ret);
            return ret;
        }


        public int GetInt32(int i)
        {
            int ret = default;
            ElementAt(i)?.TryGetInt(out ret);
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

        public string GetString(int i)
        {
            return ElementAt(i)?.Text ?? "";
        }

        public bool NextResult()
        {
            currentSheetIndex++;
            //currentSheet = sheetParts?.Skip(currentSheetIndex).FirstOrDefault()?.Worksheet.Descendants<SheetData>().First();
            rows = sheetParts.ElementAt(currentSheetIndex).Worksheet.Descendants<Row>().ToArray();
            if (headerRowIndex < rows.Length)
            {
                heads = rows[headerRowIndex].Elements<Cell>().Select(c => c.CellValue?.Text ?? "").ToArray();
            }
            return currentSheet != null;
        }

        public bool Read()
        {
            if (currentRowIndex < rows.Length)
            {
                cells = rows[currentRowIndex].Descendants<Cell>().ToArray();
                currentRowIndex++;
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
                    Close();
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
