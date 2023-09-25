using LightExcel.OpenXml;
using LightExcel.Utils;
using System.Data;

namespace LightExcel.Renders
{
	internal class DataReaderRender : RenderBase
	{
		public DataReaderRender(ExcelConfiguration configuration) : base(configuration) { }
		public override IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data)
		{
			if (data is IDataReader d)
			{
				for (int i = 0; i < d.FieldCount; i++)
				{
					var name = d.GetName(i);
					var col = new ExcelColumnInfo(name);
					col.NumberFormat = Configuration.CheckCellNumberFormat(name);
					col.Type = d.GetFieldType(i);
					col.ColumnIndex = i + 1;
					AssignDynamicInfo(col);
					yield return col;
				}
			}
		}

		public override IEnumerable<Row> RenderBody(object data, Sheet sheet, IEnumerable<ExcelColumnInfo> columns, TransConfiguration configuration)
		{
			var reader = data as IDataReader ?? throw new ArgumentException();
			var rowIndex = Configuration.StartRowIndex;
			var maxColumnIndex = 0;
			while (reader.Read())
			{
				var row = new Row() { RowIndex = ++rowIndex };
				var cellIndex = 0;
				foreach (var col in columns)
				{
					if (col.Ignore) continue;
					var value = reader.GetValue(col.ColumnIndex - 1);
					cellIndex = col.ColumnIndex;
					//var nf = configuration.NumberFormatColumnFilter(col);
					//var cell = CellHelper.CreateCell(cellIndex, rowIndex, value, col, nf, Configuration);
					var cell = new Cell();
					cell.Reference = ReferenceHelper.ConvertXyToCellReference(cellIndex, rowIndex);
					cell.Type = CellHelper.ConvertCellType(col.Type);
					cell.Value = CellHelper.GetCellValue(col, value, Configuration);
					cell.StyleIndex = col.NumberFormat || configuration.NumberFormatColumnFilter(col) ? "1" : null;
					row.AppendChild(cell);
				}
				maxColumnIndex = Math.Max(maxColumnIndex, cellIndex);
				yield return row;
			}
			sheet.MaxColumnIndex = maxColumnIndex;
			sheet.MaxRowIndex = rowIndex;
		}
	}
}