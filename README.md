# LightExcel
## 提供Excel导出功能
```CSharp
// 实例化ExcelHelper对象，调用WriteExcel
void WriteExcel(string path, object data, string sheetName = "sheet", bool appendSheet = true);

// 依赖注入
services.AddLightExcel()
```
## data参数支持`DataTable`、`DataSet`、`IDataReader`、`IEnumerable<T>`
