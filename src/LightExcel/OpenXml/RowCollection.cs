using System.Collections;
using LightExcel.OpenXml.Basic;
using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;

namespace LightExcel.OpenXml
{
    internal static class RowExtensions
    {
        public static void AddAndFixed(this Row row, Cell child)
        {
            var (x, y) = ReferenceHelper.ConvertCellReferenceToXY(child.Reference);
            while (row.Children.Count < x - 1 && x.HasValue && y.HasValue)
            {
                row.Children.Add(Cell.EmptyCell(row.Children.Count + 1, y.Value));
            }
            row.Children.Add(child);
        }
    }

    internal class Row : SimpleNodeCollectionXmlPart<Cell>
    {
        public int RowIndex { get; set; }
        public bool IsTemplateRow { get; set; }

        public override void WriteToXml(LightExcelStreamWriter writer)
        {
            if (IsTemplateRow)
            {
                return;
            }

            writer.Write($"<row r=\"{RowIndex}\">");
            foreach (Cell cell in Children)
            {
                cell.WriteToXml(writer);
            }
            writer.Write("</row>");
        }
    }
}