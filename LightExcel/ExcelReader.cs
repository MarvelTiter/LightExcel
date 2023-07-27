using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using LightExcel.Extensions;
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
        //private readonly IEnumerable<WorksheetPart> sheetParts;
        private readonly IEnumerable<Sheet> sheets;
        private readonly SpreadsheetDocument document;
        public string? this[int i] => CellAt(i)?.GetCellValue(document.WorkbookPart);
        public string? this[string name] => CellAt(GetOrdinal(name))?.GetCellValue(document.WorkbookPart);

        Row[] rows = Array.Empty<Row>();

        public string CurrentSheetName => sheets.ElementAt(currentSheetIndex)?.Name?.ToString() ?? "";

        public int FieldCount => cells.Length;

        Cell[] cells = Array.Empty<Cell>();
        string[] heads = Array.Empty<string>();
        int currentSheetIndex = 0;
        int currentRowIndex = 0;
        int startRowIndex = 0;
        public ExcelReader(SpreadsheetDocument document)
        {
            this.document = document;
            //sheetParts = document.WorkbookPart?.WorksheetParts ?? Enumerable.Empty<WorksheetPart>();
            sheets = document.WorkbookPart?.Workbook.Descendants<Sheet>() ?? Enumerable.Empty<Sheet>();
        }

        private CellValue? ElementAt(int i)
        {
            if (i > cells.Length) return null;
            return cells[i].CellValue;
        }

        private Cell? CellAt(int i)
        {
            if (i > cells.Length) return null;
            return cells[i];
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
            //rows = sheetParts.ElementAt(currentSheetIndex).Worksheet.Descendants<Row>().ToArray();
            if (currentSheetIndex >= sheets.Count()) return false;
            var sheet = sheets.ElementAt(currentSheetIndex);
            var current = (WorksheetPart)document.WorkbookPart?.GetPartById(sheet!.Id!)!;
            rows = current.Worksheet.Descendants<Row>().ToArray();
            if (rows.Length > 0)
            {
                heads = rows[0].Descendants<Cell>().Select(c => c.GetCellValue(document.WorkbookPart) ?? "").ToArray();
            }
            currentRowIndex = 0;
            return rows.Length > 0;
        }

        public bool Read()
        {
            if (currentRowIndex < rows.Length)
            {
                cells = GetCells(rows[currentRowIndex]).ToArray();
                currentRowIndex++;
                return true;
            }
            return false;
        }

        IEnumerable<Cell> GetCells(Row row)
        {
            foreach (Cell item in row)
            {
                yield return item;
            }
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
