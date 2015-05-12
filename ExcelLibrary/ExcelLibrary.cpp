// ExcelLibrary.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"
#include <comutil.h>
#include <iostream>
#include <atlconv.h>

#include "ExcelApplication.h"
#include "ExcelWorkbook.h"
#include "ExcelWorksheet.h"
#include "ExcelData.h"
#include "UseTime.h"

void Create(PCTSTR filename)
{
	ExcelApplication app;
	ExcelWorkbook workbook = app.AddWorkbook();
	workbook.SaveAs(filename, ExcelFileFormat::Xlsx);
	workbook.Close();
}

void SaveData(PCTSTR filename)
{
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
}

void ReadData(PCTSTR filename)
{ 
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
}

int _tmain(int argc, _TCHAR* argv[])
{
	CoInitialize(NULL);

	{
		UseTime useTime;
		PCTSTR filename = _T("b:\\test");

		Create(filename);
		SaveData(filename);
		ReadData(filename);
	}

	CoUninitialize();
	return 0;
}

