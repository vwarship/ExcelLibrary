#pragma once

class ExcelWorksheet
{
public:
	ExcelWorksheet(_WorksheetPtr worksheet);

	void SetValues();

	void GetValues();

private:
	_WorksheetPtr _worksheet;

};
