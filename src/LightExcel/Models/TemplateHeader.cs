using LightExcel.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.Models;

internal readonly record struct TemplateHeader(List<Row> Rows, Sheet Sheet);

