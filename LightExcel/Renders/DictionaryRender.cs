using DocumentFormat.OpenXml.Packaging;
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
        private readonly WorkbookPart workbookPart;
        private readonly ExcelConfiguration configuration;

        public DictionaryRender(WorkbookPart workbookPart, ExcelConfiguration configuration)
        {
            this.workbookPart = workbookPart;
            this.configuration = configuration;
        }

        public IEnumerable<Row> RenderBody(object data)
        {
            var values = (IEnumerable<Dictionary<string, object>>)data;
            foreach (var dic in values)
            {
                var row = new Row();
                foreach (var kv in dic)
                {
                    var cell = InternalHelper.CreateTypedCell(kv.Value.GetType(), kv.Value);
                    if (configuration.HasStyle(kv.Key, kv.Value))
                    {
                        cell.StyleIndex = configuration.GetStyleIndex(kv.Key, workbookPart);
                    }
                    else
                    {
                        cell.StyleIndex = null;
                    }
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
