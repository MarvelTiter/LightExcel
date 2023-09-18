using LightExcel.OpenXml.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.OpenXml.Styles
{
    internal class FontCollection : INodeCollection<Font>
    {
        public int Count => throw new NotImplementedException();

        public void AppendChild(Font child)
        {
            throw new NotImplementedException();
        }
    }

    internal class Font : INode
    {
        public void WriteToXml(LightExcelStreamWriter writer)
        {
            throw new NotImplementedException();
        }
    }
}
