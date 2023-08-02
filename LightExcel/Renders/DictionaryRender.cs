using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.Renders
{
    internal class DictionaryRender : IDataRender
    {
        public IEnumerable<Row> RenderBody(object data)
        {
            var values = (IEnumerable<Dictionary<string, object>>)data;
            int rowValueIndex = 0;
            foreach (var dic in values)
            {
                var row = new Row();
                foreach (var kv in dic)
                {
                    var cell = InternalHelper.CreateTypedCell(kv.Value.GetType(), kv.Value);
                    cell.CellReference = $"{kv.Key}{rowValueIndex}";
                    row.AppendChild(cell);
                }
                rowValueIndex++;
                yield return row;
            }
        }

        public Row RenderHeader(object data)
        {
            var values = (IEnumerable<Dictionary<string, object>>)data;
            var row = new Row();
            foreach (var item in values.First().Keys)
            {
                var cell = new Cell
                {
                    CellValue = new CellValue(item),
                    DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String),
                    CellReference = $"Header{item}"
                };
                row.AppendChild(cell);
            }
            return row;
        }
    }
}
