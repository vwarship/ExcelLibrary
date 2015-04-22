#pragma once
#include "ExcelWorkbook.h"

class ExcelApplication
{
public:
	ExcelApplication();

	ExcelWorkbook AddWorkbook();

	ExcelWorkbook Open(PCTSTR filename);

	~ExcelApplication();

private:
	_ApplicationPtr _app;

};

