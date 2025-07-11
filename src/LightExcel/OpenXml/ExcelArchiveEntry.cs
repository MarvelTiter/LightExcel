﻿using System.IO.Compression;
using System.Text;

namespace LightExcel.OpenXml
{
	internal class ExcelArchiveEntry : IDisposable
	{
		readonly ZipArchive archive;
		private readonly Stream stream;
		private readonly ExcelConfiguration configuration;
		private bool disposedValue;
		internal readonly static UTF8Encoding Utf8WithBom = new(true);

		public ExcelArchiveEntry(Stream stream, ExcelConfiguration configuration)
		{
			this.stream = stream;
			this.configuration = configuration;
			archive = new ZipArchive(stream, ZipArchiveMode.Update, true, Utf8WithBom);
			WorkBook = new WorkBook(archive, this, configuration);
			ContentTypes = new ContentTypes(archive);
		}
		internal WorkBook WorkBook { get; set; }
		//internal RelationshipCollection Relationship { get; set; } = new RelationshipCollection();
		internal ContentTypes ContentTypes { get; set; }
		internal void AddWorkBook()
		{
			ContentTypes.AppendChild(new Override("xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"));
		}
		internal void AddEntry(string path, string contentType, string content)
		{
			var zipEntry = archive.CreateEntry(path, CompressionLevel.Fastest);
			using var entryStream = zipEntry.Open();
			using var writer = new LightExcelStreamWriter(entryStream, Utf8WithBom, 1024 * 512);
			writer.Write(content);
			if (!string.IsNullOrEmpty(contentType))
				ContentTypes.AppendChild(new Override(path, contentType));
		}


		internal void Save()
		{
			if (configuration.Readonly) return;
			if (configuration.FillByTemplate) return;
			WorkBook.Save();
			ContentTypes.Write();
		}

		protected virtual void Dispose(bool disposing)
		{
			if (!disposedValue)
			{
				if (disposing)
				{
					archive?.Dispose();
					stream?.Dispose();
					WorkBook.Dispose();
				}
				disposedValue = true;
			}
		}


		public void Dispose()
		{
			// 不要更改此代码。请将清理代码放入“Dispose(bool disposing)”方法中
			Dispose(disposing: true);
			GC.SuppressFinalize(this);
		}
	}
}
