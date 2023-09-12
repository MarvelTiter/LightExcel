using DocumentFormat.OpenXml.Spreadsheet;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace LightExcel.OpenXml
{
    internal class Row
    {
        public int RowIndex { get; set; }
        public List<Cell> RowDatas { get; set; } = new List<Cell> ();
    }
    internal class RowCollection : XmlPart<Row>
    {
        private readonly string path;

        public RowCollection(ZipArchive archive, string path) : base(archive)
        {
            this.path = path;
            LoadStream(path);
        }

        protected override IEnumerable<Row> GetChildren()
        {
            if (reader == null) { yield break; }
            while (reader.Read())
            {
                if (reader.Name == "row")
                {
                    var row = new Row();
                    while (reader.Read())
                    {
                        if (reader.Name == "c")
                        {
                            var c = new Cell
                            {
                                Reference = reader["r"],
                                Type = reader["t"],
                                StyleIndex = reader["s"],
                            };
                            reader.Read();
                            c.Value = reader.ReadInnerXml();
                            row.RowDatas.Add(c);
                        }
                        else if (reader.Name == "row") { break; }
                    }
                    yield return row;
                }
            }
        }

        internal override void Save()
        {
            throw new NotImplementedException();
        }
    }
}
