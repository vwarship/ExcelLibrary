#include "stdafx.h"
#include "ExcelWorksheet.h"

ExcelWorksheet::ExcelWorksheet(_WorksheetPtr worksheet)
: _worksheet(worksheet)
{
}

void ExcelWorksheet::SetValues()
{
	SAFEARRAYBOUND bound[2];
	bound[0].lLbound = 1; bound[0].cElements = 2000;
	bound[1].lLbound = 1; bound[1].cElements = 10;
	SAFEARRAY* psa = SafeArrayCreate(VT_VARIANT, 2, bound);

	_variant_t value;
	value.vt = VT_ARRAY | VT_VARIANT;
	value.parray = psa;

	for (int r = 1; r <= 2000; ++r)
	{
		for (int c = 1; c <= 10; ++c)
		{
			LONG indices[] = { r, c };
			_variant_t v(_T("hello"));
			SafeArrayPutElement(value.parray, indices, &v);
		}
	}

	RangePtr usedRange = _worksheet->Range["A1"];
	usedRange = usedRange->GetResize(2000, 10);
	usedRange->Value2 = value;



	//SAFEARRAYBOUND bound[2];
	//bound[0].lLbound = 1; bound[0].cElements = 2000;
	//bound[1].lLbound = 1; bound[1].cElements = 10;

	//SAFEARRAY* psaData = SafeArrayCreate(VT_VARIANT, 2, bound);
	//VARIANT* pData = NULL;
	//HRESULT hr = SafeArrayAccessData(psaData, (void **)&pData);
	//if (SUCCEEDED(hr))
	//{
	//	for (int i = 0; i<10*2000; ++i, ++pData)
	//	{
	//		//(_variant_t)pData = _variant_t(_T("hello"));
	//		::VariantInit(pData);
	//		pData->vt = VT_BSTR;
	//		pData->bstrVal = SysAllocString(_T("hello"));
	//	}

	//	SafeArrayUnaccessData(psaData);
	//}
	//_variant_t value;
	//value.vt = VT_ARRAY | VT_VARIANT;
	//value.parray = psaData;

	//RangePtr usedRange = _worksheet->Range["A1"];
	//usedRange = usedRange->GetResize(2000, 10);
	//usedRange->Value2 = value;
}

void ExcelWorksheet::GetValues()
{
	RangePtr usedRange = _worksheet->UsedRange;
	_variant_t value = usedRange->Value2;

	if (value.vt == (VT_ARRAY | VT_VARIANT) &&
		::SafeArrayGetDim(value.parray) == 2)
	{
		LONG rowLBound = 0, rowUBound = 0;
		LONG colLBound = 0, colUBound = 0;

		SafeArrayGetLBound(value.parray, 1, &rowLBound);
		SafeArrayGetUBound(value.parray, 1, &rowUBound);
		SafeArrayGetLBound(value.parray, 2, &colLBound);
		SafeArrayGetUBound(value.parray, 2, &colUBound);

		for (int row = rowLBound; row <= rowUBound; ++row)
		{
			for (int col = colLBound; col <= colUBound; ++col)
			{
				LONG indices[] = { row, col };
				_variant_t v;
				SafeArrayGetElement(value.parray, indices, &v);
			}
		}
	}
}
