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

void ExcelWorkbook::Save()
{
	_workbook->Save();
}

/*
xlOpenXMLWorkbook	xlsx
xlWorkbookNormal	xls
*/
void ExcelWorkbook::SaveAs(PCTSTR filename, XlFileFormat fileFormat)
{
	variant_t vtFileName(filename);

	_workbook->SaveAs(vtFileName, fileFormat, vtMissing,
		vtMissing, vtMissing, vtMissing, xlNoChange);
}

void ExcelWorkbook::Close()
{
	_workbook->Close();
}
