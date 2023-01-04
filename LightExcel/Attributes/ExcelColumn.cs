using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.Attributes
{
    public class ExcelColumnAttribute : Attribute
    {
        public string? Name { get; set; }
    }
}
