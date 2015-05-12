#include "stdafx.h"
#include "ExcelData.h"

ExcelData::ExcelData()
{
	Clear();
}

void ExcelData::Create(ULONG row, ULONG col)
{
	values.Clear();

	saBound[0].lLbound = 1; saBound[0].cElements = row;
	saBound[1].lLbound = 1; saBound[1].cElements = col;
	SAFEARRAY* psa = SafeArrayCreate(VT_VARIANT, 2, saBound);

	values.vt = VT_ARRAY | VT_VARIANT;
	values.parray = psa;
}

void ExcelData::Clear()
{
	values.Clear();

	ZeroMemory(saBound, sizeof(saBound));
}

void ExcelData::SetValue(ULONG row, ULONG col, _variant_t& value)
{
	LONG indices[] = { row, col };
	SafeArrayPutElement(values.parray, indices, &value);
}

void ExcelData::GetValue(ULONG row, ULONG col, _variant_t& value) const
{
	LONG indices[] = { row, col };
	SafeArrayGetElement(values.parray, indices, &value);
}

const _variant_t& ExcelData::Values() const
{
	return values;
}

ULONG ExcelData::GetRowNum() const
{
	return saBound[0].cElements;
}

ULONG ExcelData::GetColNum() const
{
	return saBound[1].cElements;
}
