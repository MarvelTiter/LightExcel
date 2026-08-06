using System.Xml;

namespace LightExcel.OpenXml;

//internal static class LightExcelXmlReaderExtensions
//{
//    public static bool IsStartWith(this LightExcelXmlReader reader, string elementName, params string[] xmlns)
//    {
//        return xmlns.Any(ns => reader.Reader.IsStartElement(elementName, ns));
//    }
//}
internal class LightExcelXmlReader : IDisposable
{

    public LightExcelXmlReader(Stream stream, string path, ExcelConfiguration configuration)
    {
        this.stream = stream;
        Path = path;
        var setting = new XmlReaderSettings()
        {
            Async = configuration.EnableAsync
        };
        this.reader = XmlReader.Create(stream, setting);
    }

    private bool disposedValue;
    private readonly Stream stream;
    private XmlReader reader;

    public bool EOF => reader.EOF;
    public string Path { get; }
    public XmlReader Reader => reader;
    public string? this[string name] => GetAttribute(name);
    public string? this[string name, string ns] => reader.GetAttribute(name, ns);
    public string? GetAttribute(string name) => reader.GetAttribute(name);
    public string? GetAttribute(string name, string ns) => reader.GetAttribute(name, ns);

    #region 同步API
    public string ReadElementContentAsString() => reader.ReadElementContentAsString();
    public bool IsStartWith(string elementName, params string[] xmlns)
    {
        return xmlns.Any(ns => reader.IsStartElement(elementName, ns));
    }

    public bool ReadFirstContent()
    {
        if (reader.IsEmptyElement)
        {
            reader.Read();
            return false;
        }
        reader.MoveToContent();
        reader.Read();
        return true;
    }

    public void SkipNextSibling()
    {
        while (!reader.EOF)
        {
            if (!SkipContent())
                break;
        }
    }

    public bool SkipContent()
    {
        if (reader.NodeType == XmlNodeType.EndElement)
        {
            reader.Read();
            return false;
        }

        reader.Skip();
        return true;
    }
    #endregion

    #region 异步API

    public Task<string> ReadElementContentAsStringAsync() => reader.ReadElementContentAsStringAsync();


    public async Task<bool> ReadFirstContentAsync()
    {
        if (reader.IsEmptyElement)
        {
            await reader.ReadAsync();
            return false;
        }
        await reader.MoveToContentAsync();
        await reader.ReadAsync();
        return true;
    }

    public async Task SkipNextSiblingAsync()
    {
        while (!reader.EOF)
        {
            if (!await SkipContentAsync())
                break;
        }
    }

    public async Task<bool> SkipContentAsync()
    {
        if (reader.NodeType == XmlNodeType.EndElement)
        {
            await reader.ReadAsync();
            return false;
        }
        await reader.SkipAsync();
        return true;
    }

    #endregion


    protected virtual void Dispose(bool disposing)
    {
        if (!disposedValue)
        {
            if (disposing)
            {
                reader?.Dispose();
                stream?.Dispose();
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
