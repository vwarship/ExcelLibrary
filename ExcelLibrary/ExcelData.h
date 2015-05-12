#pragma once

class ExcelData
{
public:
	ExcelData();

	void Create(ULONG row, ULONG col);

	void Clear();

	void SetValue(ULONG row, ULONG col, _variant_t& value);

	void GetValue(ULONG row, ULONG col, _variant_t& value) const;

	const _variant_t& Values() const;

	ULONG GetRowNum() const;

	ULONG GetColNum() const;

private:
	SAFEARRAYBOUND saBound[2];
	_variant_t values;

};
