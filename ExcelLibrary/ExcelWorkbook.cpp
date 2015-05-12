#include "stdafx.h"
#include "ExcelWorkbook.h"

ExcelWorkbook::ExcelWorkbook(_WorkbookPtr workbook)
	: _workbook(workbook)
{
}

ExcelWorksheet ExcelWorkbook::AddWorksheet(PCTSTR name)
{
	_WorksheetPtr worksheet;

	try
	{
		SheetsPtr worksheets = _workbook->Worksheets;
		worksheet = worksheets->Add();
		worksheet->Name = name;
	}
	catch (_com_error &err)
	{
		wprintf(L"Excel throws the error: %s\n", err.ErrorMessage());
		wprintf(L"Description: %s\n", (LPCTSTR)err.Description());
	}

	return ExcelWorksheet(worksheet);
}

ExcelWorksheet ExcelWorkbook::GetWorksheet(PCTSTR name)
{
	_WorksheetPtr worksheet;

	try
	{
		SheetsPtr worksheets = _workbook->Worksheets;
		worksheet = worksheets->Item[name];
	}
	catch (_com_error &err)
	{
		wprintf(L"Excel throws the error: %s\n", err.ErrorMessage());
		wprintf(L"Description: %s\n", (LPCTSTR)err.Description());
	}

	return ExcelWorksheet(worksheet);
}

size_t ExcelWorkbook::WorksheetCount() const
{
	try
	{
		return _workbook->Worksheets->Count;
	}
	catch (_com_error &err)
	{
		wprintf(L"Excel throws the error: %s\n", err.ErrorMessage());
		wprintf(L"Description: %s\n", (LPCTSTR)err.Description());
	}

	return 0;
}

void ExcelWorkbook::Save()
{
	try
	{
		_workbook->Save();
	}
	catch (_com_error &err)
	{
		wprintf(L"Excel throws the error: %s\n", err.ErrorMessage());
		wprintf(L"Description: %s\n", (LPCTSTR)err.Description());
	}
}

/*
	xlOpenXMLWorkbook	xlsx
	xlWorkbookNormal	xls
*/
void ExcelWorkbook::SaveAs(PCTSTR filename, ExcelFileFormat fileFormat)
{
	variant_t vtFileName(filename);

	XlFileFormat xlFileFormat = xlOpenXMLWorkbook;
	if (fileFormat == ExcelFileFormat::Xls)
		xlFileFormat = xlWorkbookNormal;

	try
	{
		_workbook->SaveAs(vtFileName, xlFileFormat, vtMissing,
		vtMissing, vtMissing, vtMissing, xlNoChange);
	}
	catch (_com_error &err)
	{
		wprintf(L"Excel throws the error: %s\n", err.ErrorMessage());
		wprintf(L"Description: %s\n", (LPCTSTR)err.Description());
	}
}

void ExcelWorkbook::Close()
{
	_workbook->Close();
}
