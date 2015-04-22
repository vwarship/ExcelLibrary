#include "stdafx.h"
#include "ExcelApplication.h"

ExcelApplication::ExcelApplication()
{
	HRESULT hr = _app.CreateInstance(__uuidof(Excel::Application));
	if (FAILED(hr))
	{
		wprintf(L"CreateInstance failed w/err 0x%08lx\n", hr);
		return;
	}

	// Make Excel invisible. (i.e. Application.Visible = 0)
	_app->Visible[0] = VARIANT_FALSE;
	_app->PutDisplayAlerts(0, VARIANT_FALSE);
}

ExcelWorkbook ExcelApplication::AddWorkbook()
{
	WorkbooksPtr workbooks = _app->Workbooks;
	_WorkbookPtr workbook = workbooks->Add();

	return ExcelWorkbook(workbook);
}

ExcelWorkbook ExcelApplication::Open(PCTSTR filename)
{
	WorkbooksPtr workbooks = _app->Workbooks;
	_WorkbookPtr workbook = workbooks->Open(filename);

	return ExcelWorkbook(workbook);
}

ExcelApplication::~ExcelApplication()
{
	_app->PutDisplayAlerts(0, VARIANT_TRUE);
	_app->Quit();
}
