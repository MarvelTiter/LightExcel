using System.Data;

namespace LightExcel;

public partial class ExcelHeperExtensions
{
    public static DataSet Fill(this IExcelHelper helper, string path, string? sheetName = null, Action<ExcelConfiguration>? config = null)
    {
        using var reader = helper.ReadExcel(path, sheetName, config);
        DataSet ds = new();
        while (reader.NextResult())
        {
            var table = new DataTable(reader.CurrentSheetName);
            for (int i = 0; i < reader.FieldCount; i++)
            {
                table.Columns.Add(reader.GetName(i));
            }
            while (reader.Read())
            {
                var row = table.NewRow();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    row[i] = reader.GetValue(i);
                }
                table.Rows.Add(row);
            }
            ds.Tables.Add(table);
        }
        return ds;
    }

    public static DataSet Fill(this IExcelHelper helper, Stream stream, string? sheetName = null, Action<ExcelConfiguration>? config = null)
    {
        using var reader = helper.ReadExcel(stream, sheetName, config);
        DataSet ds = new();
        while (reader.NextResult())
        {
            var table = new DataTable(reader.CurrentSheetName);
            for (int i = 0; i < reader.FieldCount; i++)
            {
                table.Columns.Add(reader.GetName(i));
            }
            while (reader.Read())
            {
                var row = table.NewRow();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    row[i] = reader.GetValue(i);
                }
                table.Rows.Add(row);
            }
            ds.Tables.Add(table);
        }
        return ds;
    }

#if NET8_0_OR_GREATER
    public static async Task<DataSet> FillAsync(this IExcelHelper helper, string path, string? sheetName = null, Action<ExcelConfiguration>? config = null, CancellationToken cancellationToken = default)
    {
        void final(ExcelConfiguration c)
        {
            config?.Invoke(c);
            c.EnableAsync = true;
        }
        using var reader = helper.ReadExcel(path, sheetName, final);
        DataSet ds = new();
        while (await reader.NextResultAsync())
        {
            if (cancellationToken.IsCancellationRequested)
            {
                break;
            }
            var table = new DataTable(reader.CurrentSheetName);
            for (int i = 0; i < reader.FieldCount; i++)
            {
                table.Columns.Add(reader.GetName(i));
            }
            while (await reader.ReadAsync())
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    break;
                }
                var row = table.NewRow();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    row[i] = reader.GetValue(i);
                }
                table.Rows.Add(row);
            }
            ds.Tables.Add(table);
        }
        return ds;
    }

    public static async Task<DataSet> FillAsync(this IExcelHelper helper, Stream stream, string? sheetName = null, Action<ExcelConfiguration>? config = null, CancellationToken cancellationToken = default)
    {
        using var reader = helper.ReadExcel(stream, sheetName, config);
        DataSet ds = new();
        while (await reader.NextResultAsync())
        {
            if (cancellationToken.IsCancellationRequested)
            {
                break;
            }
            var table = new DataTable(reader.CurrentSheetName);
            for (int i = 0; i < reader.FieldCount; i++)
            {
                table.Columns.Add(reader.GetName(i));
            }
            while (await reader.ReadAsync())
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    break;
                }
                var row = table.NewRow();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    row[i] = reader.GetValue(i);
                }
                table.Rows.Add(row);
            }
            ds.Tables.Add(table);
        }
        return ds;
    }
#endif
}
