using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.OpenXml
{
    internal class LightExcelStreamWriter : IDisposable
    {
        private readonly Stream stream;
        private readonly Encoding encoding;
        internal readonly StreamWriter streamWriter;
        private bool disposedValue;
        public LightExcelStreamWriter(Stream stream, Encoding encoding, int bufferSize)
        {
            this.stream = stream;
            this.encoding = encoding;
            streamWriter = new StreamWriter(stream, encoding, bufferSize);
        }
        public void Write(string content)
        {
            if (string.IsNullOrEmpty(content))
                return;
            streamWriter.Write(content);
        }

        public long WriteAndFlush(string content)
        {
            Write(content);
            streamWriter.Flush();
            return streamWriter.BaseStream.Position;
        }

        public void SetPosition(long position)
        {
            streamWriter.BaseStream.Position = position;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                streamWriter?.Dispose();
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
