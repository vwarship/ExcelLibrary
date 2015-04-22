// ExcelLibrary.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"
#include <comutil.h>
#include <iostream>
#include <atlconv.h>

#include "ExcelApplication.h"
#include "ExcelWorkbook.h"
#include "ExcelWorksheet.h"


class UseTime
{
public:
	UseTime()
	{
		beginTime = ::GetTickCount();
	}

	~UseTime()
	{
		DWORD endTime = ::GetTickCount();
		std::cout << endTime - beginTime << std::endl;
	}

private:
	DWORD beginTime;

};

int _tmain(int argc, _TCHAR* argv[])
{
	CoInitialize(NULL);

	DWORD begin = ::GetTickCount();

	{
		PCTSTR filename = _T("e:\\test");

		ExcelApplication app;
		//ExcelWorkbook workbook = app.AddWorkbook();
		//workbook.SaveAs(filename, xlOpenXMLWorkbook);
		//workbook.Close();

		ExcelWorkbook workbook = app.Open(filename);
		//ExcelWorksheet worksheet = workbook.AddWorksheet(_T("测试"));
		ExcelWorksheet worksheet = workbook.GetWorksheet(_T("Sheet1"));
		DWORD _begin = ::GetTickCount();
		worksheet.GetValues();
		printf("Save: %d\n", ::GetTickCount() - _begin);
		//workbook.Save();
		workbook.Close();
	}

	printf("%d\n", ::GetTickCount() - begin);

	CoUninitialize();
	return 0;
}

