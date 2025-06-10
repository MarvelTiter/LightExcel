using LightExcel.OpenXml.Basic;
using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace LightExcel.OpenXml
{
    /// <summary>
    /// xl/sharedStrings.xml
    /// </summary>
    internal class SharedStringTable : NodeCollectionXmlPart<SharedStringNode>
    {
        public int RefCount { get; set; }
        public int UniqueCount { get; set; }
        //private readonly IDictionary<int, string> values = new Dictionary<int, string>();
        private readonly SharedStringsDiskCache diskCache;
        public override int Count => diskCache.Count;
        public SharedStringTable(ZipArchive archive) : base(archive, "xl/sharedStrings.xml")
        {
            diskCache = new SharedStringsDiskCache();
            Flush();
        }

        internal string? this[int index]
        {
            get
            {
                if (diskCache is null || index < 0 || index >= diskCache.Count)
                    return null;
                //return Children[index].Content;

                return diskCache[index];
            }
        }

        private void Flush()
        {
            // _ = GetChildren().ToList();
            //Children.Clear();
            //foreach (var s in GetChildren())
            //{
            //    Children.Add(s);
            //}
            int index = 0;
            foreach (var s in GetChildren())
            {
                diskCache.Add(index++, s.Content);
            }
        }

        public override IEnumerator<SharedStringNode> GetEnumerator() => Children.GetEnumerator();

        protected override IEnumerable<SharedStringNode> GetChildrenImpl()
        {
            if (reader is null) yield break;
            if (!reader.IsStartWith("sst", XmlHelper.MainNs)) yield break;
            if (!reader.ReadFirstContent()) yield break;
            _ = int.TryParse(reader["count"], out var count);
            _ = int.TryParse(reader["uniqueCount"], out var uniqueCount);
            RefCount = count;
            UniqueCount = uniqueCount;
            while (!reader.EOF)
            {
                if (reader.IsStartWith("si", XmlHelper.MainNs))
                {
                    var content = reader.ReadStringContent();
                    yield return new SharedStringNode(content);
                }
                else if (!reader.SkipContent())
                {
                    break;
                }
            }
        }

        protected override void WriteImpl<TNode>(LightExcelStreamWriter writer, IEnumerable<TNode> children)
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            writer.Write($"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{RefCount}\" uniqueCount=\"{UniqueCount}\">");
            foreach (var child in children)
            {
                child.WriteToXml(writer);
                //writer.Write($"<si><t>{Content}</t></si>");
            }

            writer.Write("</sst>");
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            diskCache?.Dispose();
        }
    }

    /// <summary>
    /// Copied and modified from MiniExcel - @MIT License
    /// </summary>
    internal class SharedStringsDiskCache : IDisposable
    {
        private readonly FileStream _positionFs;
        private readonly FileStream _lengthFs;
        private readonly FileStream _valueFs;
        private bool _disposedValue;
        private readonly static Encoding _encoding = new UTF8Encoding(true);
        public int Count => checked((int)_maxIndx);
        public string this[int key] { get => GetValue(key); }
        private long _maxIndx = -1;

        public SharedStringsDiskCache()
        {
            var path = $"{Guid.NewGuid().ToString()}_lightexcelcache";
            _positionFs = new FileStream($"{path}_position", FileMode.OpenOrCreate);
            _lengthFs = new FileStream($"{path}_length", FileMode.OpenOrCreate);
            _valueFs = new FileStream($"{path}_data", FileMode.OpenOrCreate);
        }

        // index must start with 0-N
        internal void Add(int index, string value)
        {
            if (index > _maxIndx)
                _maxIndx = index;
            byte[] valueBs = _encoding.GetBytes(value);
            _positionFs.Write(BitConverter.GetBytes(_valueFs.Position), 0, 4);
            _lengthFs.Write(BitConverter.GetBytes(valueBs.Length), 0, 4);
            _valueFs.Write(valueBs, 0, valueBs.Length);
        }

        private string GetValue(int index)
        {
            _positionFs.Position = index * 4;
            var bytes = new byte[4];
            _positionFs.Read(bytes, 0, 4);
            var position = BitConverter.ToInt32(bytes, 0);
            _lengthFs.Position = index * 4;
            _lengthFs.Read(bytes, 0, 4);
            var length = BitConverter.ToInt32(bytes, 0);
            _valueFs.Position = position;
            bytes = new byte[length];
            _valueFs.Read(bytes, 0, length);
            var v = _encoding.GetString(bytes);
            return v;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }
                _positionFs.Dispose();
                if (File.Exists(_positionFs.Name))
                    File.Delete(_positionFs.Name);
                _lengthFs.Dispose();
                if (File.Exists(_lengthFs.Name))
                    File.Delete(_lengthFs.Name);
                _valueFs.Dispose();
                if (File.Exists(_valueFs.Name))
                    File.Delete(_valueFs.Name);
                _disposedValue = true;
            }
        }

        ~SharedStringsDiskCache()
        {
            Dispose(disposing: false);
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}