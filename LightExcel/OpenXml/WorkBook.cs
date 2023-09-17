using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace LightExcel.OpenXml
{
    internal class WorkBook
    {
        private readonly ZipArchive archive;
        private readonly ExcelArchiveEntry doc;
        private readonly ExcelHelperConfiguration configuration;
        public WorkBook(ZipArchive archive, ExcelArchiveEntry doc, ExcelHelperConfiguration configuration)
        {
            this.archive = archive;
            this.doc = doc;
            this.configuration = configuration;
            WorkSheets = new SheetCollection(archive, configuration);
            Relationships = new RelationshipCollection(archive);
        }
        internal void Save()
        {
            Relationships.Write();
            WorkSheets.Write();
            SharedStrings?.Write();
        }
        /// <summary>
        /// xl/workbook.xml
        /// </summary>
        internal SheetCollection WorkSheets { get; set; }
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
        internal RelationshipCollection Relationships { get; set; }

        internal void AddSharedStringTable()
        {
            SharedStrings = new SharedStringTable(archive);
            Relationships.AppendChild(new Relationship($"R{Guid.NewGuid():N}", "sharedStrings", "sharedStrings.xml"));
        }

        internal void AddStyleSheet()
        {
            StyleSheet = new StyleSheet();
            Relationships.AppendChild(new Relationship($"R{Guid.NewGuid():N}", "styles", "styles.xml"));
        }

        internal Sheet AddNewSheet(string? sheetName = null)
        {
            var c = WorkSheets.Count;
            sheetName ??= $"sheet{c + 1}";
            var sheet = new Sheet(archive!, sheetName, c + 1);
            WorkSheets.AppendChild(sheet);
            Relationships.AppendChild(new Relationship(sheet.Id, "worksheet", sheet.RelPath));
            doc.ContentTypes.AppendChild(new Override(sheetName, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
            return sheet;
        }
    }
}
