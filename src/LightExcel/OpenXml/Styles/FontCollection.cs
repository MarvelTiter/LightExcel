using LightExcel.OpenXml.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LightExcel.OpenXml.Basic;

namespace LightExcel.OpenXml.Styles
{
    internal class FontCollection : SimpleNodeCollectionXmlPart<Font>
    {
       
    }

    internal class Font : INode
    {
        public void WriteToXml(LightExcelStreamWriter writer)
        {
            throw new NotImplementedException();
        }
    }
}
