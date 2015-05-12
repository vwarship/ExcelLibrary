#pragma once

class ExcelData;

class ExcelWorksheet
{
public:
	ExcelWorksheet(_WorksheetPtr worksheet);

	void SetValues(const ExcelData& excelData, const char* cell = "A1");

	void GetValues(ExcelData& excelData, const char* cell = "A1");

private:
	_WorksheetPtr _worksheet;

};
