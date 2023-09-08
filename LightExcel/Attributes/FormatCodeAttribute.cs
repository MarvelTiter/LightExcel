using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.Attributes
{
    [AttributeUsage(AttributeTargets.Field)]
    internal class FormatCodeAttribute : Attribute
    {
        public FormatCodeAttribute(string code)
        {
            Code = code;
        }

        public string Code { get; }
    }
}
