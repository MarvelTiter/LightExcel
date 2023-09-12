using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace LightExcel.OpenXml
{
    /// <summary>
    /// xl/sharedStrings.xml
    /// </summary>
    internal class SharedStringTable : XmlPart<string>
    {
        public SharedStringTable(ZipArchive archive) : base(archive)
        {

        }

        internal override void LoadStream(string path)
        {
            base.LoadStream("xl/sharedStrings.xml");
        }
        protected override IEnumerable<string> GetChildren()
        {
            throw new NotImplementedException();
        }

        internal override void Save()
        {
            throw new NotImplementedException();
        }
    }
    internal class WorkBook
    {
        private readonly ZipArchive archive;
        private readonly ExcelHelperConfiguration configuration;
        public WorkBook(ZipArchive archive, ExcelHelperConfiguration configuration)
        {
            this.archive = archive;
            this.configuration = configuration;
            WorkSheets = new SheetCollection(archive, configuration);
            Relationships = new RelationshipCollection(archive);
        }
        internal void Save()
        {
            WorkSheets.Save();
            Relationships?.Save();
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
            Relationships.Children!.Add(new Relationship($"{Guid.NewGuid():N}", "sharedStrings", "sharedStrings.xml"));
        }

        internal void AddStyleSheet()
        {
            StyleSheet = new StyleSheet();
            Relationships.Children!.Add(new Relationship($"{Guid.NewGuid():N}", "styles", "styles.xml"));
        }

        internal Sheet AddNewSheet(string sheetName)
        {
            var c = WorkSheets.Children!.Count;
            var sheet = new Sheet(archive!, sheetName, c + 1);
            WorkSheets.Children.Add(sheet);
            Relationships.Children!.Add(new Relationship(sheet.Id, "worksheet", sheet.RelPath));
            return sheet;
        }
    }
}
