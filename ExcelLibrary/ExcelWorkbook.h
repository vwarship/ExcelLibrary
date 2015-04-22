#pragma once
#include "ExcelWorksheet.h"

class ExcelWorkbook
{
public:
	ExcelWorkbook(_WorkbookPtr workbook);

	ExcelWorksheet AddWorksheet(PCTSTR name);

	ExcelWorksheet GetWorksheet(PCTSTR name);

	void Save();

	/*
	xlOpenXMLWorkbook	xlsx
	xlWorkbookNormal	xls
	*/
	void SaveAs(PCTSTR filename, XlFileFormat fileFormat);

	void Close();

private:
	_WorkbookPtr _workbook;

};
