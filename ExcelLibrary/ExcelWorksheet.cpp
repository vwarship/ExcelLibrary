#include "stdafx.h"
#include "ExcelWorksheet.h"
#include "ExcelData.h"

ExcelWorksheet::ExcelWorksheet(_WorksheetPtr worksheet)
	: _worksheet(worksheet)
{
}

void ExcelWorksheet::SetValues(const ExcelData& excelData, const char* cell/* = "A1"*/)
{
	RangePtr usedRange = _worksheet->Range[cell];
	usedRange = usedRange->GetResize(excelData.GetRowNum(), excelData.GetColNum());
	usedRange->Value2 = excelData.Values();

	//单个访问速度慢
	//RangePtr usedRange = _worksheet->Cells;
	//for (ULONG row = 1; row <= rowNum; ++row)
	//{
	//	for (ULONG col = 1; col <= colNum; ++col)
	//	{
	//		_variant_t value(".....");
	//		usedRange->Item[row][col] = value;
	//	}
	//}
}

void ExcelWorksheet::GetValues(ExcelData& excelData, const char* cell/* = "A1"*/)
{
	RangePtr usedRange = _worksheet->UsedRange;
	_variant_t values = usedRange->Value2;

	if (values.vt == (VT_ARRAY | VT_VARIANT) &&
		::SafeArrayGetDim(values.parray) == 2)
	{
		LONG rowLBound = 0, rowUBound = 0;
		LONG colLBound = 0, colUBound = 0;

		SafeArrayGetLBound(values.parray, 1, &rowLBound);
		SafeArrayGetUBound(values.parray, 1, &rowUBound);
		SafeArrayGetLBound(values.parray, 2, &colLBound);
		SafeArrayGetUBound(values.parray, 2, &colUBound);

		excelData.Create(rowUBound - rowLBound + 1, colUBound - colLBound + 1);

		for (int row = rowLBound; row <= rowUBound; ++row)
		{
			for (int col = colLBound; col <= colUBound; ++col)
			{
				LONG indices[] = { row, col };
				_variant_t value;
				SafeArrayGetElement(values.parray, indices, &value);

				excelData.SetValue(row - rowLBound + 1, col - colLBound + 1, value);
			}
		}

		//VARIANT* pData = NULL;
		//HRESULT hr = SafeArrayAccessData(values.parray, (void **)&pData);
		//if (SUCCEEDED(hr))
		//{
		//	int rowNum = rowUBound - rowLBound + 1;
		//	int colNum = colUBound - colLBound + 1;
		//	for (int row = rowLBound; row <= rowUBound; ++row)
		//	{
		//		for (int col = colLBound; col <= colUBound; ++col)
		//		{
		//			_variant_t v = pData[(col - colLBound)*rowNum + (row - rowLBound)];
		//			printf("%d\t", (int)(V_R8(&v)));
		//		}

		//		printf("\n");
		//	}

		//	SafeArrayUnaccessData(values.parray);
		//}
	}
}
