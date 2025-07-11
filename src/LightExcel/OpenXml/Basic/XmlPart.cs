﻿using System.Collections;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using LightExcel.OpenXml.Interfaces;
using LightExcel.Utils;

namespace LightExcel.OpenXml
{
    internal abstract partial class XmlPart<T> : IXmlPart<T> where T : INode
    {
        private bool disposedValue;
        protected ZipArchive? archive;
        protected LightExcelXmlReader? reader = null;
        internal virtual string Path { get; }
        public XmlPart(ZipArchive archive, string path)
        {
            this.archive = archive;
            Path = path;
        }
        protected virtual void SetXmlReader()
        {
            reader ??= archive!.GetXmlReader(Path);
        }

        public virtual void Write()
        {

        }

        public void Write<TNode>(IEnumerable<TNode> children) 
            where TNode : INode
        {
            using var writer = archive!.GetWriter(Path);
            WriteImpl(writer, children);
        }
#if NET6_0_OR_GREATER
        public async Task WriteAsync<TNode>(IAsyncEnumerable<TNode> children, CancellationToken cancellationToken = default)
            where TNode : INode
        {
            using var writer = archive!.GetWriter(Path);
            await WriteAsyncImpl(writer, children, cancellationToken);
        }
        protected virtual Task WriteAsyncImpl<TNode>(LightExcelStreamWriter writer, IAsyncEnumerable<TNode> children, CancellationToken cancellationToken = default) where TNode : INode
        {
            throw new NotSupportedException();
        }
#endif
        public void Replace<TNode>(IEnumerable<TNode> children) where TNode : INode
        {
            if (reader?.Path == Path)
            {
                reader?.Dispose();
                reader = null;
                archive!.GetEntry(Path)?.Delete();
            }
            Write(children);
        }

        public void DeleteEntry()
        {
            if (reader?.Path == Path)
            {
                reader?.Dispose();
                reader = null;
                archive!.GetEntry(Path)?.Delete();
            }
        }

        protected abstract void WriteImpl<TNode>(LightExcelStreamWriter writer, IEnumerable<TNode> children) where TNode : INode;

        

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    reader?.Dispose();
                    reader = null;
                }
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
