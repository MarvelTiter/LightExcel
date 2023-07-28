﻿using System.Data;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace LightExcel
{
    public partial class ExcelHelper
    {
        void InternalWriteExcelWithTemplate(string path, string template, object data)
        {
            var doc = GetDocument(template);

            // 获取WorkbookPart（工作簿）
            var workBookPart = doc.WorkbookPart;
            var sheets = workBookPart?.Workbook.Descendants<Sheet>().ToArray() ?? Array.Empty<Sheet>();
            var dataType = data.GetType();
            var render = RenderProvider.GetDataRender(dataType);
            if (dataType == typeof(DataSet))
            {
                if (data is DataSet ds)
                {
                    for (int i = 0; i < ds.Tables.Count; i++)
                    {
                        if (i > sheets.Length) break;
                        var sheet = sheets[i];
                        AppendData(workBookPart!, sheet);
                    }
                }
            }
            else
            {
                var sheet = sheets.First();
                AppendData(workBookPart!, sheet);
            }
            doc?.SaveAs(path);
            doc?.Dispose();

            void AppendData(WorkbookPart bookPart, Sheet sheet)
            {
                var sheetPart = (WorksheetPart)bookPart.GetPartById(sheet!.Id!);
                var headRows = sheetPart.Worksheet.Descendants<Row>().Count();
                //创建内容数据
                CreateBody(sheetPart, data, render, (uint)headRows);

                sheetPart.Worksheet.Save();
                bookPart.Workbook.Save();
            }
        }
    }
    public partial class ExcelHelper : IExcelHelper
    {
        private readonly ExcelHelperConfiguration configuration = new ExcelHelperConfiguration();
        const string DEFAULT_SHEETNAME = "sheet";
        public void WriteExcel(string path, object data, string? sheetName = "sheet", bool appendSheet = true)
        {
            try
            {
                configuration.AllowAppendSheet = appendSheet;
                InternalWriteExcel(path, data, sheetName);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void WriteExcel(string path, string template, object data)
        {
            try
            {
                configuration.AllowAppendSheet = true;
                InternalWriteExcelWithTemplate(path, template, data);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public IExcelDataReader ReadExcel(string path)
        {
            configuration.AllowAppendSheet = false;
            var doc = GetDocument(path);
            return new ExcelReader(doc);
        }

        public IEnumerable<T> QueryExcel<T>(string path, string? sheetName = null, int startRow = 2)
        {
            var reader = ReadExcel(path);
            while (reader.NextResult())
            {
                if (sheetName != null && reader.CurrentSheetName == sheetName)
                {
                    while (reader.Read())
                    {
                        //
                    }
                }
                else
                {

                }
            }
            throw new NotImplementedException();
        }

        public IEnumerable<dynamic> QueryExcel(string path, string? sheetName = null, int startRow = 2)
        {
            throw new NotImplementedException();
        }

        private void InternalWriteExcel(string path, object data, string? sheetName)
        {
            using var doc = GetDocument(path);
            var dataType = data.GetType();
            var render = RenderProvider.GetDataRender(dataType);
            if (dataType == typeof(DataSet))
            {
                configuration.AllowAppendSheet = true;
                foreach (DataTable dataTable in (data as DataSet ?? new DataSet()).Tables)
                {
                    WriteSheet(doc, data, render, dataTable.TableName);
                }
            }
            else
                WriteSheet(doc, data, render, sheetName);
        }

        private SpreadsheetDocument GetDocument(string path)
        {
            SpreadsheetDocument doc = null;
            try
            {
                if (File.Exists(path))
                {
                    // 文件存在并且，允许追加Sheet
                    using var fs = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                    MemoryStream ms = new MemoryStream();
                    fs.CopyTo(ms);
                    doc = SpreadsheetDocument.Open(ms, configuration.AllowAppendSheet);
                }
                else
                {
                    File.Delete(path);
                    //创建Workbook, 指定为Excel Workbook (*.xlsx).
                    doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
                    //创建WorkbookPart（工作簿）
                    WorkbookPart workbookPart = doc.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    //创建工作表列表
                    workbookPart.Workbook.AppendChild(new Sheets());
                    //构建SharedStringTablePart
                    var shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                    //创建共享字符串表
                    shareStringPart.SharedStringTable = new SharedStringTable();
                    workbookPart.Workbook.Save();
                }
            }
            catch (Exception)
            {
                throw;
            }
            return doc;
        }

        private void WriteSheet(SpreadsheetDocument doc, object data, IDataRender render, string? sheetName)
        {
            // 获取WorkbookPart（工作簿）
            var workBookPart = doc.WorkbookPart;
            var sheets = workBookPart!.Workbook.Sheets;
            //获取SharedStringTablePart
            var sharedStringTable = workBookPart!.SharedStringTablePart;
            //创建WorksheetPart（工作簿中的工作表）
            var worksheetPart = workBookPart!.AddNewPart<WorksheetPart>();
            var newSheetIndex = sheets!.Count() + 1;
            sheetName = sheetName == DEFAULT_SHEETNAME ? $"sheet{newSheetIndex}" : sheetName;
            //Workbook 下创建Sheets节点, 建立一个子节点Sheet，关联工作表WorksheetPart
            var rid = workBookPart.GetIdOfPart(worksheetPart);
            workBookPart!.Workbook.Sheets!.AppendChild(new Sheet()
            {
                Id = rid,
                SheetId = (uint)newSheetIndex,
                Name = sheetName
            });

            //初始化Worksheet
            InitWorksheet(worksheetPart);

            //创建表头
            CreateHeader(worksheetPart, sharedStringTable!, data, render);

            //创建内容数据
            CreateBody(worksheetPart, data, render);

            worksheetPart.Worksheet.Save();
            workBookPart.Workbook.Save();
        }

        /// <summary>
        /// 初始化工作表
        /// </summary>
        /// <param name="worksheetPart"></param>
        private void InitWorksheet(WorksheetPart worksheetPart)
        {
            var worksheet = new Worksheet();
            //SheetFormatProperties, 设置默认行高度，宽度， 值类型是Double类型。
            var sheetFormatProperties = new SheetFormatProperties()
            {
                DefaultColumnWidth = 15d,
                DefaultRowHeight = 15d
            };
            // 顺序不能变
            worksheet.Append(new OpenXmlElement[]
            {
                sheetFormatProperties,
                new Columns(),
                new SheetData()
            });
            worksheetPart.Worksheet = worksheet;
        }

        /// <summary>
        /// 创建表头
        /// </summary>
        /// <param name="worksheetPart">WorksheetPart 对象</param>
        /// <param name="shareStringPart">SharedStringTablePart 对象</param>
        private void CreateHeader(WorksheetPart worksheetPart, SharedStringTablePart shareStringPart, object data, IDataRender render)
        {
            //获取Worksheet对象
            var worksheet = worksheetPart.Worksheet;

            //获取表格的数据对象，SheetData
            var sheetData = worksheet.GetFirstChild<SheetData>();

            var row = render.RenderHeader(data);
            row.RowIndex = 1;
            sheetData!.AppendChild(row);
        }

        private void CreateBody(WorksheetPart worksheetPart, object data, IDataRender render, uint rowIndex = 2)
        {
            //获取Worksheet对象
            var worksheet = worksheetPart.Worksheet;

            //获取表格的数据对象，SheetData
            var sheetData = worksheet.GetFirstChild<SheetData>();

            var rows = render.RenderBody(data);
            uint startIndex = rowIndex;
            foreach (var r in rows)
            {
                r.RowIndex = startIndex;
                sheetData!.AppendChild(r);
                startIndex++;
            }
        }


    }
}
