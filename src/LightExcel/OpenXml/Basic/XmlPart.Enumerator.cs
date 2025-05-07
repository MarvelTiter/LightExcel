using LightExcel.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace LightExcel.OpenXml
{
    internal abstract partial class XmlPart<T>
    {
        protected virtual LightExcelXmlReader? GetXmlReader()
        {
            return archive!.GetXmlReader(Path);
        }
        protected virtual IEnumerable<T> GetChildren()
        {
            using var reader = GetXmlReader();
            if (reader == null) yield break;
            cached ??= new List<T>();
            foreach (var item in GetChildrenImpl(reader))
            {
                cached.Add(item);
                yield return item;
            }
        }

        protected abstract IEnumerable<T> GetChildrenImpl(LightExcelXmlReader reader);

        public IEnumerator<T> GetEnumerator()
        {
            if (cached == null)
            {
                return GetChildren().GetEnumerator();
            }
            else
            {
                return cached.GetEnumerator();
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
