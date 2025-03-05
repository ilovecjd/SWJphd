// XLAutomation.h: interface for the CXLAutomation class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_XLAUTOMATION_H__E020CE95_7428_4BEF_A24C_48CE9323C450__INCLUDED_)
#define AFX_XLAUTOMATION_H__E020CE95_7428_4BEF_A24C_48CE9323C450__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <vector>

// CXLAutomation class definition
class CXLAutomation
{

#define MAX_DISP_ARGS 10
#define DISPARG_NOFREEVARIANT 0x01
#define DISP_FREEARGS 0x02
#define DISP_NOSHOWEXCEPTIONS 0x03
#define xlWorksheet -4167
#define xl3DPie -4102
#define xlRows 1
#define xlXYScatter -4169
#define xlXYScatterLines 74
#define xlXYScatterSmoothNoMarkers 73
#define xlXYScatterSmooth 72
#define xlXYScatterLinesNoMarkers 75
#define xlColumns 2
#define xlNormal -4143
#define xlUp -4162

#define xlContinuous 1 // xlContinuous for borders line style
#define xlThin 2       // xlThin for border weight
#define xlAutomatic -4105 // xlAutomatic for borders color

public:
	//BOOL OpenExcelFile(CString szFileName);	
	BOOL OpenExcelFile(CString szFileName, LPOLESTR sheetsName[], int nSheetCount);
	BOOL OpenExcelFile(CString szFileName, CString strSheetName);
	BOOL SaveAs(CString szFileName, int nFileFormat, CString szPassword, CString szWritePassword, BOOL bReadOnly, BOOL bBackUp);
	BOOL DeleteRow(int sheet, long nRow);
	BOOL ReleaseExcel();

	BOOL AddArgumentCStringArray(LPOLESTR lpszArgName, WORD wFlags, LPOLESTR *paszStrings, int iCount);
	BOOL SetRangeValueDouble(int sheet, LPOLESTR lpszRef, double d);
	BOOL SetCellsValueToString(int sheet, double Row, double Column, CString szStr);
	BOOL AddArgumentOLEString(LPOLESTR lpszArgName, WORD wFlags, LPOLESTR lpsz);
	BOOL AddArgumentCString(LPOLESTR lpszArgName, WORD wFlags, CString szStr);
	//BOOL CreateWorkSheet(LPOLESTR strSheetName);
	BOOL AddArgumentDouble(LPOLESTR lpszArgName, WORD wFlags, double d);
	BOOL AddArgumentBool(LPOLESTR lpszArgName, WORD wFlags, BOOL b);
	BOOL AddArgumentInt2(LPOLESTR lpszArgName, WORD wFlags, int i);
	BOOL AddArgumentDispatch(LPOLESTR lpszArgName, WORD wFlags, IDispatch* pdisp);
	void AddArgumentCommon(LPOLESTR lpszArgName, WORD wFlags, VARTYPE vt);
	BOOL InitOLE();

	//song 
	BOOL SetRangeValueAndStyle(int sheet, int startRow, int startCol, int** dataArray, int numRows, int numCols);
	BOOL GetRange(int sheet, int startRow, int startCol, int endRow, int endCol, VARIANTARG* pRange);
	BOOL ReadRangeToArray(int sheet, int startRow, int startCol, int endRow, int endCol, int* dataArray, int rows, int cols);
	
	BOOL GetCellValueInt(int sheet, int nRow, int nColumn, int* pValue);
	BOOL GetCellValueCString(int sheet, int nRow, int nColumn, CString* pValue);
	BOOL GetCellValueDouble(int sheet, int nRow, int nColumn, double* pValue);
	BOOL GetCellValueVariant(int sheet, int nRow, int nColumn, VARIANTARG* pValue); // 범용 함수 선언

	// 셀에 값을 설정하는 함수들 (오버로딩)
	BOOL SetCellValueInt(int sheet, int nRow, int nColumn, int value);
	BOOL SetCellValueCString(int sheet, int nRow, int nColumn, CString value);
	BOOL SetCellValueDouble(int sheet, int nRow, int nColumn, double value);
	
	// For reading values from Excel
	BOOL ReadRangeToIntArray(int sheet, int startRow, int startCol, int* dataArray, int rows, int cols);
	BOOL ReadRangeToCStringArray(int sheet, int startRow, int startCol, CString* dataArray, int rows, int cols);

	BOOL WriteArrayToRangeInt(int sheet, int startRow, int startCol, int* dataArray, int rows, int cols);
	BOOL WriteArrayToRangeCString(int sheet, int startRow, int startCol, CString* dataArray, int rows, int cols);
	BOOL WriteArrayToRangeVariant(int sheet, int startRow, int startCol, VARIANT* dataArray, int rows, int cols);

	HRESULT AutoWrap(int autoType, VARIANT* pvResult, IDispatch* pDisp, LPOLESTR ptName, int cArgs...);

	BOOL SetRangeBorder(int sheet, int startRow, int startCol, int endRow, int endCol, int borderStyle, int borderWeight, int borderColor);
	BOOL SetRangeBorderAround(int sheet, int startRow, int startCol, int endRow, int endCol, int borderStyle, int borderWeight, int borderColor);

	BOOL ReadExRangeConvertInt(int sheet, int startRow, int startCol, int* dataArray, int rows, int cols);

	BOOL AddNewSheet(CString sheetName);

	BOOL SaveAndCloseExcelFile(CString szFileName);
	// 자동화 헬퍼 함수
	//song

	CXLAutomation();
	CXLAutomation(BOOL bVisible);
	virtual ~CXLAutomation();

protected:
	void ShowException(LPOLESTR szMember, HRESULT hr, EXCEPINFO *pexcep, unsigned int uiArgErr);
	void ReleaseDispatch();
	BOOL SetExcelVisible(BOOL bVisible);
	void ReleaseVariant(VARIANTARG *pvarg);
	void ClearAllArgs();
	void ClearVariant(VARIANTARG *pvarg);

	int m_iArgCount;
	int m_iNamedArgCount;
	VARIANTARG m_aVargs[MAX_DISP_ARGS];
	DISPID m_aDispIds[MAX_DISP_ARGS + 1]; // one extra for the member name
	LPOLESTR m_alpszArgNames[MAX_DISP_ARGS + 1]; // used to hold the argnames for GetIDs
	WORD m_awFlags[MAX_DISP_ARGS];

	BOOL ExlInvoke(IDispatch* pdisp, LPOLESTR szMember, VARIANTARG* pvargReturn,
		WORD wInvokeAction, WORD wFlags);
	IDispatch* m_pdispExcelApp;
	IDispatch* m_pdispWorkbook;
	// 기존: IDispatch* m_pdispWorksheets[WS_TOTAL_SHEET_COUNT];// Array to store worksheet dispatch interfaces
	std::vector<IDispatch*> m_pdispWorksheets; 

	BOOL StartExcel();
	BOOL CXLAutomation::FindAndStoreWorksheet(IDispatch* pWorkbook, LPOLESTR sheetName);
};

#endif // !defined(AFX_XLAUTOMATION_H__E020CE95_7428_4BEF_A24C_48CE9323C450__INCLUDED_)
