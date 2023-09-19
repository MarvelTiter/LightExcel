# LightExcel

## 基于 OpenXml 的 Excel 导出功能

### 简单导出 Excel 文件

```CSharp
// 实例化ExcelHelper对象，调用WriteExcel
void WriteExcel(string path, object data, Action<ExcelConfiguration>? config);
```

### 简单模板导出 Excel 文件

```CSharp
// 基于简单模板的Excel导出
var ie = Datas.DictionarySource();
ExcelHelper excel = new ExcelHelper();
// 目标文件，模板文件，数据源
excel.WriteExcelByTemplate("12test.xlsx", "路檢報表格式.xlsx", ie, config: config =>
{

});
```

### 在同一个 Excel 文件中追加 sheet

```CSharp
// 支持分段写入Excel文件，在trans对象释放前，会在同一个excel文件中追加sheet
using var trans = excel.BeginTransaction("1test.xlsx");
trans.WriteExcel(data);
```

### 依赖注入使用，获取`IExcelHelper`对象

```CSharp
// 依赖注入
services.AddLightExcel()
```

## data 参数支持`DataTable`、`IDataReader`、`IDictionary<string, object>`、`IEnumerable<T>`
