#pragma once
#include "ExcelWorksheet.h"

enum ExcelFileFormat
{
	Xls,
	Xlsx
};

class ExcelWorkbook
{
public:
	ExcelWorkbook(_WorkbookPtr workbook);

	ExcelWorksheet AddWorksheet(PCTSTR name);

	ExcelWorksheet GetWorksheet(PCTSTR name);

	size_t WorksheetCount() const;

	void Save();

	void SaveAs(PCTSTR filename, ExcelFileFormat fileFormat);

	void Close();

private:
	_WorkbookPtr _workbook;

};
