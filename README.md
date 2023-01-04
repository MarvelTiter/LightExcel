# LightExcel
## 基于OpenXml的Excel导出功能
```CSharp
// 实例化ExcelHelper对象，调用WriteExcel
void WriteExcel(string path, object data, string sheetName = "sheet", bool appendSheet = true);

// 依赖注入
services.AddLightExcel()
```
## data参数支持`DataTable`、`DataSet`、`IDataReader`、`IEnumerable<T>`
## 支持从现有的Excel文件中追加sheet
