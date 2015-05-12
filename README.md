## Excel Library

## 介绍
包装了 Excel 的 COM 接口。快速读取和写入大量数据。

## 功能

* 创建 xls, xlsx 格式

```C++
ExcelApplication app;
ExcelWorkbook workbook = app.AddWorkbook();
workbook.SaveAs(filename, ExcelFileFormat::Xlsx);
workbook.Close();
```

* 写入数据

```C++
const ULONG rowNum = 20;
const ULONG colNum = 10;
ExcelData excelData;
excelData.Create(rowNum, colNum);

for (ULONG row = 1; row <= rowNum; ++row)
{
	for (ULONG col = 1; col <= colNum; ++col)
	{
		_variant_t value((row - 1)*colNum + col);
		excelData.SetValue(row, col, value);
	}
}

ExcelApplication app;
ExcelWorkbook workbook = app.Open(filename);
ExcelWorksheet worksheet = workbook.GetWorksheet(_T("Sheet1"));
worksheet.SetValues(excelData);
workbook.Save();
workbook.Close();
```

* 读取数据

```C++
ExcelData excelData;

ExcelApplication app;
ExcelWorkbook workbook = app.Open(filename);
ExcelWorksheet worksheet = workbook.GetWorksheet(_T("Sheet1"));
worksheet.GetValues(excelData);
workbook.Close();

for (ULONG row = 1; row <= excelData.GetRowNum(); ++row)
{
	for (ULONG col = 1; col <= excelData.GetColNum(); ++col)
	{
		_variant_t value;
		excelData.GetValue(row, col, value);
		std::cout << V_R8(&value) << '\t';
	}
	std::cout << std::endl;
}
```
