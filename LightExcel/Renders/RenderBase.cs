﻿using LightExcel.OpenXml;
using LightExcel.Utils;
using System.Data;
using System.Reflection.PortableExecutable;

namespace LightExcel.Renders
{
    internal abstract class RenderBase : IDataRender
    {
        public ExcelConfiguration Configuration { get; }
        public RenderBase(ExcelConfiguration configuration)
        {
            Configuration = configuration;
        }


        public abstract IEnumerable<ExcelColumnInfo> CollectExcelColumnInfo(object data);

        public abstract IEnumerable<Row> RenderBody(object data, Sheet sheet, IEnumerable<ExcelColumnInfo> columns, TransConfiguration configuration);

        public virtual Row RenderHeader(IEnumerable<ExcelColumnInfo> columns, TransConfiguration configuration)
        {
            var row = new Row() { RowIndex = 1 };
            Configuration.StartRowIndex = 1;
            var index = 0;
            foreach (var col in columns)
            {
                var cell = new Cell
                {
                    Reference = ReferenceHelper.ConvertXyToCellReference(++index, 1),
                    Type = "str",
                    Value = col.Name
                };
                row.AppendChild(cell);
            }
            return row;
        }
    }
}