using LightExcel.OpenXml.Interfaces;

namespace LightExcel.OpenXml.Styles
{
    internal class BorderCollection : INodeCollection<Border>, INode
    {
        internal IList<Border> Borders { get; set; } = new List<Border>();
        public int Count => Borders.Count;

        public void AppendChild(Border child)
        {
            Borders.Add(child);
        }

        public void WriteToXml(LightExcelStreamWriter writer)
        {
            writer.Write($"<borders count={Count}>");
            foreach (var item in Borders)
            {
                item.WriteToXml(writer);
            }
            writer.Write("</borders>");
        }
    }

    internal class Border : IStyleNode
    {
        public int Id { get; }
        public void WriteToXml(LightExcelStreamWriter writer)
        {

        }
    }
}
