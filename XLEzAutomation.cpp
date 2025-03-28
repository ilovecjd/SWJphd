﻿// EzAutomation.cpp: implementation of the CXLEzAutomation class.
//This wrapper classe is provided for easy access to basic automation  
//methods of the CXLAutoimation.
//Only very basic set of methods is provided here.
//
//////////////////////////////////////////////////////////////////////

#include "pch.h"
#include "GlobalEnv.h"
#include "XLEzAutomation.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//

CXLEzAutomation::CXLEzAutomation()
{
	//Starts Excel with bVisible = TRUE and creates empty worksheet 
	m_pXLServer = new CXLAutomation;
}

CXLEzAutomation::CXLEzAutomation(BOOL bVisible)
{
	//Can be used to start Excel in background (bVisible = FALSE)
	m_pXLServer = new CXLAutomation(bVisible);

}

CXLEzAutomation::~CXLEzAutomation()
{
	if(NULL != m_pXLServer)
		delete m_pXLServer;
}



//Quit Excel
BOOL CXLEzAutomation::ReleaseExcel()
{
	return m_pXLServer->ReleaseExcel();
}
//Delete line from worksheet
BOOL CXLEzAutomation::DeleteRow(int sheet, int nRow)
{
	return m_pXLServer->DeleteRow(sheet, nRow);
}
//Save workbook as Excel file
BOOL CXLEzAutomation::SaveFileAs(CString szFileName)
{
	return m_pXLServer->SaveAs(szFileName, xlNormal, _T(""), _T(""), FALSE, FALSE);
}

//Open Excell file
BOOL CXLEzAutomation::OpenExcelFile(CString szFileName, LPOLESTR sheetsName[], int nSheetCount)
{
	return m_pXLServer->OpenExcelFile(szFileName, sheetsName, nSheetCount);
}
BOOL CXLEzAutomation::OpenExcelFile(CString szFileName, CString szSheetName, BOOL bCreate)
{
	return m_pXLServer->OpenExcelFile(szFileName, szSheetName, bCreate);
}

// Overloaded GetCellValue functions
// Returns integer value from Worksheet.Cells(nColumn, nRow)
BOOL CXLEzAutomation::GetCellValue(int sheet, int nRow, int nColumn, int* pValue)
{
	if (pValue == nullptr) return FALSE;
	return m_pXLServer->GetCellValueInt(sheet, nRow, nColumn, pValue);
}

// Returns CString value from Worksheet.Cells(nRow, nColumn)
BOOL CXLEzAutomation::GetCellValue(int sheet, int nRow, int nColumn, CString* pValue)
{
	if (pValue == nullptr) return FALSE;
	return m_pXLServer->GetCellValueCString(sheet, nRow, nColumn, pValue);
}

// Returns double value from Worksheet.Cells(nColumn, nRow)
BOOL CXLEzAutomation::GetCellValue(int sheet, int nRow, int nColumn, double* pValue)
{
	if (pValue == nullptr) return FALSE;
	return m_pXLServer->GetCellValueDouble(sheet, nRow, nColumn, pValue);
}

// SetCellValue for integer
BOOL CXLEzAutomation::SetCellValue(int sheet, int nRow, int nColumn, int value)
{
	if (m_pXLServer == NULL)
		return FALSE;
	return m_pXLServer->SetCellValueInt(sheet, nRow, nColumn, value);
}

// SetCellValue for CString
BOOL CXLEzAutomation::SetCellValue(int sheet, int nRow, int nColumn, CString value)
{
	if (m_pXLServer == NULL)
		return FALSE;
	return m_pXLServer->SetCellValueCString(sheet, nRow, nColumn, value);
}

// SetCellValue for double
BOOL CXLEzAutomation::SetCellValue(int sheet, int nRow, int nColumn, double value)
{
	if (m_pXLServer == NULL)
		return FALSE;
	return m_pXLServer->SetCellValueDouble(sheet, nRow, nColumn, value);
}

// Overloaded function to read integer values from Excel (int)
BOOL CXLEzAutomation::ReadRangeToArray(int sheet, int startRow, int startCol, int* dataArray, int rows, int cols)
{
	if (!m_pXLServer) return FALSE;
	return m_pXLServer->ReadRangeToIntArray(sheet, startRow, startCol, dataArray, rows, cols);
}

// Overloaded function to read CString values from Excel (CSting)
BOOL CXLEzAutomation::ReadRangeToArray(int sheet, int startRow, int startCol, CString* dataArray, int rows, int cols)
{
	if (!m_pXLServer) return FALSE;
	return m_pXLServer->ReadRangeToCStringArray(sheet, startRow, startCol, dataArray, rows, cols);
}

// int 배열을 Excel에 쓰기
BOOL CXLEzAutomation::WriteArrayToRange(int sheet, int startRow, int startCol, int* dataArray, int rows, int cols)
{
	if (m_pXLServer == NULL)
		return FALSE;
	return m_pXLServer->WriteArrayToRangeInt(sheet, startRow, startCol, dataArray, rows, cols);
}

// CString 배열을 Excel에 쓰기
BOOL CXLEzAutomation::WriteArrayToRange(int sheet, int startRow, int startCol, CString* dataArray, int rows, int cols)
{
	if (m_pXLServer == NULL)
		return FALSE;
	return m_pXLServer->WriteArrayToRangeCString(sheet, startRow, startCol, dataArray, rows, cols);
}

BOOL CXLEzAutomation::WriteArrayToRange(int sheet, int startRow, int startCol, VARIANT* dataArray, int rows, int cols)
{
	if (m_pXLServer == NULL)
		return FALSE;
	return m_pXLServer->WriteArrayToRangeVariant(sheet, startRow, startCol, dataArray, rows, cols);
}

BOOL CXLEzAutomation::SetRangeBorder(int sheet, int startRow, int startCol, int endRow, int endCol, int borderStyle, int borderWeight, int borderColor)
{
	// Call the underlying CXLAutomation function to set borders
	return m_pXLServer->SetRangeBorder(sheet, startRow, startCol, endRow, endCol, borderStyle, borderWeight, borderColor);
}

BOOL CXLEzAutomation::SetRangeBorderAround(int sheet, int startRow, int startCol, int endRow, int endCol, int borderStyle, int borderWeight, int borderColor)
{	
	// Call the function from CXLAutomation
	return m_pXLServer->SetRangeBorderAround(sheet, startRow, startCol, endRow, endCol, borderStyle, borderWeight, borderColor);
}

BOOL CXLEzAutomation::ReadExRangeConvertInt(int sheet, int startRow, int startCol, int* dataArray, int rows, int cols)
{
	// Call the function from CXLAutomation
	int transRows = startRow + rows - 1;
	int transCols = startCol + cols - 1;

	return m_pXLServer->ReadExRangeConvertInt(sheet, startRow, startCol, dataArray, transRows, transCols);
}

BOOL CXLEzAutomation::SaveAndCloseExcelFile(CString szFileName)
{
	return m_pXLServer->SaveAndCloseExcelFile(szFileName);
}

BOOL CXLEzAutomation::ClearSheetContents(int sheet)
{
	return m_pXLServer->ClearSheetContents(sheet);
}

BOOL CXLEzAutomation::CreateExcelFile(CString szFileName, LPOLESTR sheetsName[], int nSheetCount)
{
	return m_pXLServer->CreateExcelFile(szFileName, sheetsName, nSheetCount);
}

BOOL CXLEzAutomation::SetCellFormula(int sheet, int nRow, int nColumn, CString formula)
{
	return m_pXLServer->SetCellFormula(sheet, nRow, nColumn, formula);
}