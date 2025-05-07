using LightExcel.OpenXml.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.OpenXml.Styles
{
    internal interface IStyleCollection<T> : INodeCollection<T> where T : INode
    {
        StyleType Type { get; }
    }
}
