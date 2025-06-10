using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace LightExcel.OpenXml
{
    internal class WorkBook(ZipArchive archive, ExcelArchiveEntry doc, ExcelConfiguration configuration) : IDisposable
    {
        private bool disposedValue;
        private RelationshipCollection? relationships;
        private SheetCollection? workSheets;

        internal void Save()
        {
            WorkSheets.Write();
            if (!configuration.FillByTemplate)
            {
                Relationships.Write();
                SharedStrings?.Write();
                StyleSheet?.Write();
            }
        }
        /// <summary>
        /// xl/workbook.xml
        /// </summary>
        internal SheetCollection WorkSheets
        {
            get
            {
                workSheets ??= new SheetCollection(archive, configuration);
                return workSheets;
            }
            set => workSheets = value;
        }
        /// <summary>
        /// xl/sharedStrings.xml
        /// </summary>
        internal SharedStringTable? SharedStrings { get; set; }
        /// <summary>
        /// xl/styles.xml
        /// </summary>
        internal StyleSheet? StyleSheet { get; set; }

        /// <summary>
        /// xl/_rels/workbook.xml.rels
        /// </summary>
        internal RelationshipCollection Relationships
        {
            get
            {
                relationships ??= new RelationshipCollection(archive);
                return relationships;
            }
            set => relationships = value;
        }
        internal void InitSharedStringTable()
        {
            SharedStrings = new SharedStringTable(archive);
        }
        internal void AddSharedStringTable()
        {
            InitSharedStringTable();
            Relationships.AppendChild(new Relationship($"R{Guid.NewGuid():N}", "sharedStrings", "sharedStrings.xml"));
            doc.ContentTypes.AppendChild(new Override("xl/sharedStrings.xml", "application/vnd.openxmlformats-package.relationships+xml"));
        }

        internal void InitStyleSheet()
        {
            StyleSheet = new StyleSheet(archive);
        }

        internal void AddStyleSheet()
        {
            InitStyleSheet();
            Relationships.AppendChild(new Relationship($"R{Guid.NewGuid():N}", "styles", "styles.xml"));
            doc.ContentTypes.AppendChild(new Override("xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"));
        }

        internal Sheet AddNewSheet(string? sheetName = null)
        {
            var c = WorkSheets.Count;
            sheetName ??= $"sheet{c + 1}";
            var sheet = new Sheet(archive!, sheetName, c + 1);
            WorkSheets.AppendChild(sheet);
            Relationships.AppendChild(new Relationship(sheet.Id, "worksheet", sheet.RelPath));
            doc.ContentTypes.AppendChild(new Override(sheet.Path, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
            return sheet;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                }
                SharedStrings?.Dispose();
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            // 不要更改此代码。请将清理代码放入“Dispose(bool disposing)”方法中
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
