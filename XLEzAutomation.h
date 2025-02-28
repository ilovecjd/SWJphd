// EzAutomation.h: interface for the CXLEzAutomation class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_EzAutomation_H__D140B9A3_1995_40AC_8E6D_8F23A95A63A2__INCLUDED_)
#define AFX_EzAutomation_H__D140B9A3_1995_40AC_8E6D_8F23A95A63A2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "XLAutomation.h"
#define xlNormal -4143

class CXLEzAutomation
{
public:
	BOOL OpenExcelFile(CString szFileName);
	BOOL OpenExcelFile(CString szFileName, CString szSheetName);
	BOOL SaveFileAs(CString szFileName);
	BOOL ReleaseExcel();

	// Overloaded GetCellValue functions
	BOOL GetCellValue(SheetName sheet, int nRow, int nColumn, int* pValue);      // For int
	BOOL GetCellValue(SheetName sheet, int nRow, int nColumn, CString* pValue);  // For CString
	BOOL GetCellValue(SheetName sheet, int nRow, int nColumn, double* pValue);   // For double

	// 셀에 값을 설정하는 함수들 (오버로딩)
	BOOL SetCellValue(SheetName sheet, int nRow, int nColumn, int value);
	BOOL SetCellValue(SheetName sheet, int nRow, int nColumn, CString value);
	BOOL SetCellValue(SheetName sheet, int nRow, int nColumn, double value);

	// Overloaded ReadRangeToArray functions
	BOOL ReadRangeToArray(SheetName sheet, int startRow, int startCol, int* dataArray, int rows, int cols);
	BOOL ReadRangeToArray(SheetName sheet, int startRow, int startCol, CString* dataArray, int rows, int cols);

	// 배열을 Excel에 쓰기
	BOOL WriteArrayToRange(SheetName sheet, int startRow, int startCol, int* dataArray, int rows, int cols);
	BOOL WriteArrayToRange(SheetName sheet, int startRow, int startCol, CString* dataArray, int rows, int cols);
	BOOL WriteArrayToRange(SheetName sheet, int startRow, int startCol, VARIANT* dataArray, int rows, int cols);

	BOOL SetRangeBorder(SheetName sheet, int startRow, int startCol, int endRow, int endCol, int borderStyle, int borderWeight, int borderColor);
	BOOL SetRangeBorderAround(SheetName sheet, int startRow, int startCol, int endRow, int endCol, int borderStyle, int borderWeight, int borderColor);
	
	BOOL ReadExRangeConvertInt(SheetName sheet, int startRow, int startCol, int* dataArray, int rows, int cols);

	BOOL DeleteRow(SheetName sheet, int nRow);

	BOOL ExportCString(SheetName sheet, CString szDataCollection);

	BOOL SaveAndCloseExcelFile(CString szFileName);

	CXLEzAutomation();
	CXLEzAutomation(BOOL bVisible);
	virtual ~CXLEzAutomation();

protected:
	CXLAutomation* m_pXLServer;
};

#endif // !defined(AFX_EzAutomation_H__D140B9A3_1995_40AC_8E6D_8F23A95A63A2__INCLUDED_)
