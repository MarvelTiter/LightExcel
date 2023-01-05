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
            foreach (var dic in values)
            {
                var row = new Row();
                foreach (var v in dic.Values)
                {
                    var cell = InternalHelper.CreateTypedCell(v.GetType(), v);
                    row.AppendChild(cell);
                }
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
                };
                row.AppendChild(cell);
            }
            return row;
        }
    }
}
