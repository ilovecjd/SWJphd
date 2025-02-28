// XLAutomation.cpp: implementation of the CXLAutomation class.
//This is C++ modification of the AutoXL C-sample from 
//Microsoft Excel97 Developer Kit, Microsoft Press 1997 
//////////////////////////////////////////////////////////////////////

#include "pch.h"
#include "GlobalEnv.h"
#include "XLAutomation.h"
#include <ole2ver.h>
#include <string.h>
#include <winuser.h>
#include <stdio.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////
/*
 *  Arrays of argument information, which are used to build up the arg list
 *  for an IDispatch call.  These arrays are statically allocated to reduce
 *  complexity, but this code could be easily modified to perform dynamic
 *  memory allocation.
 *
 *  When arguments are added they are placed into these arrays.  The
 *  Vargs array contains the argument values, and the lpszArgNames array
 *  contains the name of the arguments, or a NULL if the argument is unnamed.
 *  Flags for the argument such as NOFREEVARIANT are kept in the wFlags array.
 *
 *  When Invoke is called, the names in the lpszArgNames array are converted
 *  into the DISPIDs expected by the IDispatch::Invoke function.  The
 *  IDispatch::GetIDsOfNames function is used to perform the conversion, and
 *  the resulting IDs are placed in the DispIds array.  There is an additional
 *  slot in the DispIds and lpszArgNames arrays to allow for the name and DISPID
 *  of the method or property being invoked.
 *  
 *  Because these arrays are static, it is important to call the ClearArgs()
 *  function before setting up arguments.  ClearArgs() releases any memory
 *  in use by the argument array and resets the argument counters for a fresh
 *  Invoke.
 */
//int			m_iArgCount;
//int			m_iNamedArgCount;
//VARIANTARG	m_aVargs[MAX_DISP_ARGS];
//DISPID		m_aDispIds[MAX_DISP_ARGS + 1];		// one extra for the member name
//LPOLESTR	m_alpszArgNames[MAX_DISP_ARGS + 1];	// used to hold the argnames for GetIDs
//WORD		m_awFlags[MAX_DISP_ARGS];
//////////////////////////////////////////////////////////////////////

CXLAutomation::CXLAutomation()
{
	m_pdispExcelApp = NULL;
	m_pdispWorkbook = NULL;
	// m_pdispActiveChart = NULL;

	// Initialize worksheet pointers to NULL
	for (int i = 0; i < WS_TOTAL_SHEET_COUNT; i++)
		m_pdispWorksheets[i] = NULL;

	InitOLE();
	StartExcel();
	SetExcelVisible(TRUE);
	// CreateWorkSheet();  // Not creating an empty worksheet anymore
}

CXLAutomation::CXLAutomation(BOOL bVisible)
{
	m_pdispExcelApp = NULL;
	m_pdispWorkbook = NULL;
	// m_pdispActiveChart = NULL;

	// Initialize worksheet pointers to NULL
	for (int i = 0; i < WS_TOTAL_SHEET_COUNT; i++)
		m_pdispWorksheets[i] = NULL;

	InitOLE();
	StartExcel();
	SetExcelVisible(bVisible);
	// CreateWorkSheet();  // Not creating an empty worksheet anymore
}

CXLAutomation::~CXLAutomation()
{
	//ReleaseExcel();
	ReleaseDispatch();
	OleUninitialize();
}

BOOL CXLAutomation::InitOLE()
{
	DWORD dwOleVer;
	
	dwOleVer = CoBuildVersion();
	
	// check the OLE library version
	if (rmm != HIWORD(dwOleVer)) 
	{
		MessageBox(NULL, _T("Incorrect version of OLE libraries."), _T("Failed"), MB_OK | MB_ICONSTOP);
		return FALSE;
	}
	
	// could also check for minor version, but this application is
	// not sensitive to the minor version of OLE
	
	// initialize OLE, fail application if we can't get OLE to init.
	if (FAILED(OleInitialize(NULL))) 
	{
		MessageBox(NULL, _T("Cannot initialize OLE."), _T("Failed"), MB_OK | MB_ICONSTOP);
		return FALSE;
	}
	
		
	return TRUE;

}

BOOL CXLAutomation::StartExcel()
{
	CLSID clsExcelApp;

	// if Excel is already running, return with current instance
	if (m_pdispExcelApp != NULL)
		return TRUE;

	/* Obtain the CLSID that identifies EXCEL.APPLICATION
	 * This value is universally unique to Excel versions 5 and up, and
	 * is used by OLE to identify which server to start.  We are obtaining
	 * the CLSID from the ProgID.
	 */
	if (FAILED(CLSIDFromProgID(L"Excel.Application", &clsExcelApp))) 
	{
		MessageBox(NULL, _T("Cannot obtain CLSID from ProgID"), _T("Failed"), MB_OK | MB_ICONSTOP);
		return FALSE;
	}

	// start a new copy of Excel, grab the IDispatch interface
	if (FAILED(CoCreateInstance(clsExcelApp, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&m_pdispExcelApp))) 
	{
		MessageBox(NULL, _T("Cannot start an instance of Excel for Automation."), _T("Failed"), MB_OK | MB_ICONSTOP);
		return FALSE;
	}

	return TRUE;

}

/*******************************************************************
 *
 *								INVOKE
 *
 *******************************************************************/

/*
 *  INVOKE
 *
 *  Invokes a method or property.  Takes the IDispatch object on which to invoke,
 *  and the name of the method or property as a String.  Arguments, if any,
 *  must have been previously setup using the AddArgumentXxx() functions.
 *
 *  Returns TRUE if the call succeeded.  Returns FALSE if an error occurred.
 *  A messagebox will be displayed explaining the error unless the DISP_NOSHOWEXCEPTIONS
 *  flag is specified.  Errors can be a result of unrecognized method or property
 *  names, bad argument names, invalid types, or runtime-exceptions defined
 *  by the recipient of the Invoke.
 *
 *  The argument list is reset via ClearAllArgs() if the DISP_FREEARGS flag is
 *  specified.  If not specified, it is up to the caller to call ClearAllArgs().
 *
 *  The return value is placed in pvargReturn, which is allocated by the caller.
 *  If no return value is required, pass NULL.  It is up to the caller to free
 *  the return value (ReleaseVariant()).
 *
 *  This function calls IDispatch::GetIDsOfNames for every invoke.  This is not
 *  very efficient if the same method or property is invoked multiple times, since
 *  the DISPIDs for a particular method or property will remain the same during
 *  the lifetime of an IDispatch object.  Modifications could be made to this code
 *  to cache DISPIDs.  If the target application is always the same, a similar
 *  modification is to statically browse and store the DISPIDs at compile-time, since
 *  a given application will return the same DISPIDs in different sessions.
 *  Eliminating the extra cross-process GetIDsOfNames call can result in a
 *  signficant time savings.
 */


BOOL CXLAutomation::ExlInvoke(IDispatch *pdisp, LPOLESTR szMember, VARIANTARG * pvargReturn,
			WORD wInvokeAction, WORD wFlags)
{
	HRESULT hr;
	DISPPARAMS dispparams;
	unsigned int uiArgErr;
	EXCEPINFO excep;
	
	// Get the IDs for the member and its arguments.  GetIDsOfNames expects the
	// member name as the first name, followed by argument names (if any).
	m_alpszArgNames[0] = szMember;
	hr = pdisp->GetIDsOfNames( IID_NULL, m_alpszArgNames,
								1 + m_iNamedArgCount, LOCALE_SYSTEM_DEFAULT, m_aDispIds);
	if (FAILED(hr)) 
	{
		if (!(wFlags & DISP_NOSHOWEXCEPTIONS))
			ShowException(szMember, hr, NULL, 0);
		return FALSE;
	}
	
	if (pvargReturn != NULL)
		ClearVariant(pvargReturn);
	
	// if doing a property put(ref), we need to adjust the first argument to have a
	// named arg of DISPID_PROPERTYPUT.
	if (wInvokeAction & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) 
	{
		m_iNamedArgCount = 1;
		m_aDispIds[1] = DISPID_PROPERTYPUT;
		pvargReturn = NULL;
	}
	
	dispparams.rgdispidNamedArgs = m_aDispIds + 1;
	dispparams.rgvarg = m_aVargs;
	dispparams.cArgs = m_iArgCount;
	dispparams.cNamedArgs = m_iNamedArgCount;
	
	excep.pfnDeferredFillIn = NULL;
	
	hr = pdisp->Invoke(m_aDispIds[0], IID_NULL, LOCALE_SYSTEM_DEFAULT,
								wInvokeAction, &dispparams, pvargReturn, &excep, &uiArgErr);
	
	if (wFlags & DISP_FREEARGS)
		ClearAllArgs();
	
	if (FAILED(hr)) 
	{
		// display the exception information if appropriate:
		if (!(wFlags & DISP_NOSHOWEXCEPTIONS))
			ShowException(szMember, hr, &excep, uiArgErr);
	
		// free exception structure information
		SysFreeString(excep.bstrSource);
		SysFreeString(excep.bstrDescription);
		SysFreeString(excep.bstrHelpFile);
	
		return FALSE;
	}
	return TRUE;
}

/*
 *  ClearVariant
 *
 *  Zeros a variant structure without regard to current contents
 */
void CXLAutomation::ClearVariant(VARIANTARG *pvarg)
{
	pvarg->vt = VT_EMPTY;
	pvarg->wReserved1 = 0;
	pvarg->wReserved2 = 0;
	pvarg->wReserved3 = 0;
	pvarg->lVal = 0;

}

/*
 *  ClearAllArgs
 *
 *  Clears the existing contents of the arg array in preparation for
 *  a new invocation.  Frees argument memory if so marked.
 */
void CXLAutomation::ClearAllArgs()
{
	int i;
	
	for (i = 0; i < m_iArgCount; i++) 
	{
		if (m_awFlags[i] & DISPARG_NOFREEVARIANT)
			// free the variant's contents based on type
			ClearVariant(&m_aVargs[i]);
		else
			ReleaseVariant(&m_aVargs[i]);
	}

	m_iArgCount = 0;
	m_iNamedArgCount = 0;

}

/*
 *  ReleaseVariant
 *
 *  Clears a particular variant structure and releases any external objects
 *  or memory contained in the variant.  Supports the data types listed above.
 */
void CXLAutomation::ReleaseVariant(VARIANTARG *pvarg)
{
	VARTYPE vt;
	VARIANTARG *pvargArray;
	long lLBound, lUBound, l;
	
	vt = pvarg->vt & 0xfff;		// mask off flags
	
	// check if an array.  If so, free its contents, then the array itself.
	if (V_ISARRAY(pvarg)) 
	{
		// variant arrays are all this routine currently knows about.  Since a
		// variant can contain anything (even other arrays), call ourselves
		// recursively.
		if (vt == VT_VARIANT) 
		{
			SafeArrayGetLBound(pvarg->parray, 1, &lLBound);
			SafeArrayGetUBound(pvarg->parray, 1, &lUBound);
			
			if (lUBound > lLBound) 
			{
				lUBound -= lLBound;
				
				SafeArrayAccessData(pvarg->parray, (void**)&pvargArray);
				
				for (l = 0; l < lUBound; l++) 
				{
					ReleaseVariant(pvargArray);
					pvargArray++;
				}
				
				SafeArrayUnaccessData(pvarg->parray);
			}
		}
		else 
		{
			MessageBox(NULL, _T("ReleaseVariant: Array contains non-variant type"), _T("Failed"), MB_OK | MB_ICONSTOP);
		}
		
		// Free the array itself.
		SafeArrayDestroy(pvarg->parray);
	}
	else 
	{
		switch (vt) 
		{
			case VT_DISPATCH:
				//(*(pvarg->pdispVal->lpVtbl->Release))(pvarg->pdispVal);
				pvarg->pdispVal->Release();
				break;
				
			case VT_BSTR:
				SysFreeString(pvarg->bstrVal);
				break;
				
			case VT_I2:
			case VT_BOOL:
			case VT_R8:
			case VT_ERROR:		// to avoid erroring on an error return from Excel
				// no work for these types
				break;
				
			default:
				MessageBox(NULL, _T("ReleaseVariant: Unknown type"), _T("Failed"), MB_OK | MB_ICONSTOP);
				break;
		}
	}
	
	ClearVariant(pvarg);

}

BOOL CXLAutomation::SetExcelVisible(BOOL bVisible)
{
	if (m_pdispExcelApp == NULL)
		return FALSE;
	
	ClearAllArgs();
	AddArgumentBool(NULL, 0, bVisible);
	return ExlInvoke(m_pdispExcelApp, L"Visible", NULL, DISPATCH_PROPERTYPUT, DISP_FREEARGS);

}

/*******************************************************************
 *
 *					   ARGUMENT CONSTRUCTOR FUNCTIONS
 *
 *  Each function adds a single argument of a specific type to the list
 *  of arguments for the current invoke.  If appropriate, memory may be
 *  allocated to represent the argument.  This memory will be
 *  automatically freed the next time ClearAllArgs() is called unless
 *  the NOFREEVARIANT flag is specified for a particular argument.  If
 *  NOFREEVARIANT is specified it is the responsibility of the caller
 *  to free the memory allocated for or contained within the argument.
 *
 *  Arguments may be named.  The name string must be a C-style string
 *  and it is owned by the caller.  If dynamically allocated, the caller
 *  must free the name string.
 *
 *******************************************************************/

/*
 *  Common code used by all variant types for setting up an argument.
 */

void CXLAutomation::AddArgumentCommon(LPOLESTR lpszArgName, WORD wFlags, VARTYPE vt)
{
	ClearVariant(&m_aVargs[m_iArgCount]);
	
	m_aVargs[m_iArgCount].vt = vt;
	m_awFlags[m_iArgCount] = wFlags;
	
	if (lpszArgName != NULL) 
	{
		m_alpszArgNames[m_iNamedArgCount + 1] = lpszArgName;
		m_iNamedArgCount++;
	}
}	
	

BOOL CXLAutomation::AddArgumentDispatch(LPOLESTR lpszArgName, WORD wFlags, IDispatch * pdisp)
{
	AddArgumentCommon(lpszArgName, wFlags, VT_DISPATCH);
	m_aVargs[m_iArgCount++].pdispVal = pdisp;
	return TRUE;
}


BOOL CXLAutomation::AddArgumentInt2(LPOLESTR lpszArgName, WORD wFlags, int i)
{
	AddArgumentCommon(lpszArgName, wFlags, VT_I2);
	m_aVargs[m_iArgCount++].iVal = i;
	return TRUE;
}


BOOL CXLAutomation::AddArgumentBool(LPOLESTR lpszArgName, WORD wFlags, BOOL b)
{
	AddArgumentCommon(lpszArgName, wFlags, VT_BOOL);
	// Note the variant representation of True as -1
	m_aVargs[m_iArgCount++].boolVal = b ? -1 : 0;
	return TRUE;
}

BOOL CXLAutomation::AddArgumentDouble(LPOLESTR lpszArgName, WORD wFlags, double d)
{
	AddArgumentCommon(lpszArgName, wFlags, VT_R8);
	m_aVargs[m_iArgCount++].dblVal = d;
	return TRUE;
}


BOOL CXLAutomation::ReleaseExcel()
{
	if (m_pdispExcelApp == NULL)
		return TRUE;
	
	// Tell Excel to quit, since for automation simply releasing the IDispatch
	// object isn't enough to get the server to shut down.
	
	// Note that this code will hang if Excel tries to display any message boxes.
	// This can occur if a document is in need of saving.  The CreateChart() code
	// always clears the dirty bit on the documents it creates, avoiding this problem.
	ClearAllArgs();
	ExlInvoke(m_pdispExcelApp, L"Quit", NULL, DISPATCH_METHOD, 0);
	
	// Even though Excel has been told to Quit, we still need to release the
	// OLE object to account for all memory.
	ReleaseDispatch();
	
	return TRUE;

}


/*
 *  OLE and IDispatch use a BSTR as the representation of strings.
 *  This constructor automatically copies the passed-in C-style string
 *  into a BSTR.  It is important to not set the NOFREEVARIANT flag
 *  for this function, otherwise the allocated BSTR copy will probably
 *  get lost and cause a memory leak.
 */

BOOL CXLAutomation::AddArgumentOLEString(LPOLESTR lpszArgName, WORD wFlags, LPOLESTR lpsz)
{
	BSTR b;
	
	b = SysAllocString(lpsz);
	if (!b)
		return FALSE;
	AddArgumentCommon(lpszArgName, wFlags, VT_BSTR);
	m_aVargs[m_iArgCount++].bstrVal = b;
	return TRUE;

}

BOOL CXLAutomation::AddArgumentCString(LPOLESTR lpszArgName, WORD wFlags, CString szStr)
{
	BSTR b;
	
	b = szStr.AllocSysString();
	if (!b)
		return FALSE;
	AddArgumentCommon(lpszArgName, wFlags, VT_BSTR);
	m_aVargs[m_iArgCount++].bstrVal = b;
	
	return TRUE;
}

//Perform Worksheets.Cells(x,y).Value = szStr
BOOL CXLAutomation::SetCellsValueToString(SheetName sheet, double Row, double Column, CString szStr)
{
	if (NULL == m_pdispWorksheets[sheet])
		return FALSE;
	if (szStr.IsEmpty())
		return FALSE;

	VARIANTARG vargRng;

	ClearAllArgs();
	AddArgumentDouble(NULL, 0, Row);
	AddArgumentDouble(NULL, 0, Column);	
	if (!ExlInvoke(m_pdispWorksheets[sheet], L"Cells", &vargRng, DISPATCH_PROPERTYGET, DISP_FREEARGS))
		return FALSE;

	AddArgumentCString(NULL, 0, szStr);
	if (!ExlInvoke(vargRng.pdispVal, L"Value", NULL, DISPATCH_PROPERTYPUT, 0))
		return FALSE;
	ReleaseVariant(&vargRng);

	return TRUE;
}

BOOL CXLAutomation::SetRangeValueDouble(SheetName sheet, LPOLESTR lpszRef, double d)
{
	if (NULL == m_pdispWorksheets[sheet])
		return FALSE;

	VARIANTARG vargRng;
	BOOL fResult;

	ClearAllArgs();
	AddArgumentOLEString(NULL, 0, lpszRef);
	if (!ExlInvoke(m_pdispWorksheets[sheet], L"Range", &vargRng, DISPATCH_PROPERTYGET, DISP_FREEARGS))
		return FALSE;

	AddArgumentDouble(NULL, 0, d);
	fResult = ExlInvoke(vargRng.pdispVal, L"Value", NULL, DISPATCH_PROPERTYPUT, 0);
	ReleaseVariant(&vargRng);

	return fResult;
}

/*
 *  Constructs an 1-dimensional array containing variant strings.  The strings
 *  are copied from an incoming array of C-Strings.
 */
BOOL CXLAutomation::AddArgumentCStringArray(LPOLESTR lpszArgName, WORD wFlags, LPOLESTR *paszStrings, int iCount)
{
	SAFEARRAY *psa;
	SAFEARRAYBOUND saBound;
	VARIANTARG *pvargBase;
	VARIANTARG *pvarg;
	int i, j;
	
	saBound.lLbound = 0;
	saBound.cElements = iCount;
	
	psa = SafeArrayCreate(VT_VARIANT, 1, &saBound);
	if (psa == NULL)
		return FALSE;
	
	SafeArrayAccessData(psa, (void**) &pvargBase);
	
	pvarg = pvargBase;
	for (i = 0; i < iCount; i++) 
	{
		// copy each string in the list of strings
		ClearVariant(pvarg);
		pvarg->vt = VT_BSTR;
		if ((pvarg->bstrVal = SysAllocString(*paszStrings++)) == NULL) 
		{
			// memory failure:  back out and free strings alloc'ed up to
			// now, and then the array itself.
			pvarg = pvargBase;
			for (j = 0; j < i; j++) 
			{
				SysFreeString(pvarg->bstrVal);
				pvarg++;
			}
			SafeArrayDestroy(psa);
			return FALSE;
		}
		pvarg++;
	}
	
	SafeArrayUnaccessData(psa);

	// With all memory allocated, setup this argument
	AddArgumentCommon(lpszArgName, wFlags, VT_VARIANT | VT_ARRAY);
	m_aVargs[m_iArgCount++].parray = psa;
	return TRUE;

}



//Clean up: release dipatches
void CXLAutomation::ReleaseDispatch()
{
	if (NULL != m_pdispExcelApp) {
		m_pdispExcelApp->Release();
		m_pdispExcelApp = NULL;
	}

	if (NULL != m_pdispWorkbook) {
		m_pdispWorkbook->Release();
		m_pdispWorkbook = NULL;
	}

	// Release all worksheet dispatch pointers
	for (int i = 0; i < WS_TOTAL_SHEET_COUNT; i++) {
		if (m_pdispWorksheets[i] != NULL) {
			m_pdispWorksheets[i]->Release();
			m_pdispWorksheets[i] = NULL;
		}
	}
}

void CXLAutomation::ShowException(LPOLESTR szMember, HRESULT hr, EXCEPINFO *pexcep, unsigned int uiArgErr)
{
	TCHAR szBuf[512];
	
	switch (GetScode(hr)) 
	{
		case DISP_E_UNKNOWNNAME:
			wsprintf(szBuf, TEXT("%s: Unknown name or named argument."), szMember);
			break;
	
		case DISP_E_BADPARAMCOUNT:
			wsprintf(szBuf, TEXT("%s: Incorrect number of arguments."), szMember);
			break;
			
		case DISP_E_EXCEPTION:
			wsprintf(szBuf, TEXT("%s: Error %d: "), szMember, pexcep->wCode);
			if (pexcep->bstrDescription != NULL)
				lstrcat(szBuf, (LPCWSTR)pexcep->bstrDescription);
			else
				lstrcat(szBuf, TEXT("<<No Description>>"));
			break;
			
		case DISP_E_MEMBERNOTFOUND:
			wsprintf(szBuf, TEXT("%s: method or property not found."), szMember);
			break;
		
		case DISP_E_OVERFLOW:
			wsprintf(szBuf, TEXT("%s: Overflow while coercing argument values."), szMember);
			break;
		
		case DISP_E_NONAMEDARGS:
			wsprintf(szBuf, TEXT("%s: Object implementation does not support named arguments."),
						szMember);
		    break;
		    
		case DISP_E_UNKNOWNLCID:
			wsprintf(szBuf, TEXT("%s: The locale ID is unknown."), szMember);
			break;
		
		case DISP_E_PARAMNOTOPTIONAL:
			wsprintf(szBuf, TEXT("%s: Missing a required parameter."), szMember);
			break;
		
		case DISP_E_PARAMNOTFOUND:
			wsprintf(szBuf, TEXT("%s: Argument not found, argument %d."), szMember, uiArgErr);
			break;
			
		case DISP_E_TYPEMISMATCH:
			wsprintf(szBuf, TEXT("%s: Type mismatch, argument %d."), szMember, uiArgErr);
			break;

		default:
			wsprintf(szBuf, TEXT("%s: Unknown error occured."), szMember);
			break;
	}
	
	MessageBox(NULL, szBuf, TEXT("OLE Error"), MB_OK | MB_ICONSTOP);

}
//Delete entire line from the current worksheet
//Worksheet.Rows(nLine).Select
//Selection.Delete Shift:=xlUp
BOOL CXLAutomation::DeleteRow(SheetName sheet, long nRow)
{
	if (NULL == m_pdispWorksheets[sheet])
		return FALSE;

	VARIANTARG varg1;

	ClearAllArgs();
	AddArgumentDouble(NULL, 0, nRow);
	if (!ExlInvoke(m_pdispWorksheets[sheet], L"Rows", &varg1, DISPATCH_PROPERTYGET, DISP_FREEARGS))
		return FALSE;

	ClearAllArgs();
	AddArgumentInt2(L"Shift", 0, xlUp);
	if (!ExlInvoke(varg1.pdispVal, L"Delete", NULL, DISPATCH_METHOD, DISP_FREEARGS))
		return FALSE;

	return TRUE;
}

//Save current workbook as an Excel file:
//ActiveWorkbook.SaveAs
//FileName:=szFileName, FileFormat:=xlNormal,
//Password:=szPassword,
//WriteResPassword:=szWritePassword,
//ReadOnlyRecommended:= bReadOnly,
//CreateBackup:= bBackup
BOOL CXLAutomation::SaveAs(CString szFileName, int nFileFormat, CString szPassword, CString szWritePassword, BOOL bReadOnly, BOOL bBackUp)
{
	if(NULL == m_pdispWorkbook)
		return FALSE;
	ClearAllArgs();
	AddArgumentBool(L"CreateBackup", 0, bBackUp);
	AddArgumentBool(L"ReadOnlyRecommended", 0, bReadOnly);
	AddArgumentCString(L"WriteResPassword", 0, szWritePassword);
	AddArgumentCString(L"Password", 0, szPassword);
	AddArgumentCString(L"FileName", 0, szFileName);
	if (!ExlInvoke(m_pdispWorkbook, L"SaveAs", NULL, DISPATCH_METHOD, DISP_FREEARGS))
		return FALSE;

	return TRUE;
}

//Open Microsoft Excel file and switch to the firs available worksheet. 
BOOL CXLAutomation::OpenExcelFile(CString szFileName) {
	// Leave if the file cannot be opened
	if (NULL == m_pdispExcelApp)
		return FALSE;
	if (szFileName.IsEmpty())
		return FALSE;

	VARIANTARG varg1, vargWorkbook;
	ClearAllArgs();
	if (!ExlInvoke(m_pdispExcelApp, L"Workbooks", &varg1, DISPATCH_PROPERTYGET, 0))
		return FALSE;

	ClearAllArgs();
	AddArgumentCString(L"Filename", 0, szFileName);
	if (!ExlInvoke(varg1.pdispVal, L"Open", &vargWorkbook, DISPATCH_METHOD, DISP_FREEARGS))
		return FALSE;

	m_pdispWorkbook = vargWorkbook.pdispVal;

	// Loop through sheet names and find corresponding worksheet objects
	for (int i = 0; i < WS_TOTAL_SHEET_COUNT; i++) {
		if (!FindAndStoreWorksheet(m_pdispWorkbook, gSheetNames[i], &m_pdispWorksheets[i])) {
			// Error handling if a sheet is not found
			MessageBox(NULL, _T("Worksheet not found."), _T("Error"), MB_OK | MB_ICONSTOP);
			return FALSE;
		}
	}

	return TRUE;
}


//Open Microsoft Excel file and switch to the firs available worksheet. 
BOOL CXLAutomation::OpenExcelFile(CString szFileName, CString strSheetName) {
	// Leave if the file cannot be opened
	if (NULL == m_pdispExcelApp)
		return FALSE;
	if (szFileName.IsEmpty())
		return FALSE;

	// Check if the file exists
	CFileFind fileFinder;
	BOOL fileExists = fileFinder.FindFile(szFileName);

	VARIANTARG varg1, vargWorkbook;
	ClearAllArgs();
	if (!ExlInvoke(m_pdispExcelApp, L"Workbooks", &varg1, DISPATCH_PROPERTYGET, 0))
		return FALSE;

	if (fileExists) {
		// If the file exists, open it
		ClearAllArgs();
		AddArgumentCString(L"Filename", 0, szFileName);
		if (!ExlInvoke(varg1.pdispVal, L"Open", &vargWorkbook, DISPATCH_METHOD, DISP_FREEARGS))
			return FALSE;
	}
	else {
		// If the file doesn't exist, create a new workbook and save it
		ClearAllArgs();
		if (!ExlInvoke(varg1.pdispVal, L"Add", &vargWorkbook, DISPATCH_METHOD, DISP_FREEARGS))
			return FALSE;

		// Save the newly created workbook with the provided file name
		ClearAllArgs();
		AddArgumentCString(L"Filename", 0, szFileName);
		if (!ExlInvoke(vargWorkbook.pdispVal, L"SaveAs", NULL, DISPATCH_METHOD, DISP_FREEARGS))
			return FALSE;
	}

	m_pdispWorkbook = vargWorkbook.pdispVal;

	// Add new sheet if specified
	AddNewSheet(strSheetName);

	BSTR b;
	b = strSheetName.AllocSysString();

	// Find and store the newly added sheet
	if (!FindAndStoreWorksheet(m_pdispWorkbook, b, &m_pdispWorksheets[WS_NUM_DEBUG_INFO])) {
		// Error handling if a sheet is not found
		MessageBox(NULL, _T("Worksheet not found."), _T("Error"), MB_OK | MB_ICONSTOP);
		return FALSE;
	}

	return TRUE;
}

//BOOL CXLAutomation::OpenExcelFile(CString szFileName, CString strSheetName) {
//	// Leave if the file cannot be opened
//	if (NULL == m_pdispExcelApp)
//		return FALSE;
//	if (szFileName.IsEmpty())
//		return FALSE;
//
//	VARIANTARG varg1, vargWorkbook;
//	ClearAllArgs();
//	if (!ExlInvoke(m_pdispExcelApp, L"Workbooks", &varg1, DISPATCH_PROPERTYGET, 0))
//		return FALSE;
//
//	ClearAllArgs();
//	AddArgumentCString(L"Filename", 0, szFileName);
//	if (!ExlInvoke(varg1.pdispVal, L"Open", &vargWorkbook, DISPATCH_METHOD, DISP_FREEARGS))
//		return FALSE;
//
//	m_pdispWorkbook = vargWorkbook.pdispVal;
//
//	AddNewSheet(strSheetName);
//
//	BSTR b;
//	b = strSheetName.AllocSysString();
//
//	if (!FindAndStoreWorksheet(m_pdispWorkbook, b, &m_pdispWorksheets[WS_NUM_DEBUG_INFO])) {
//		// Error handling if a sheet is not found
//		MessageBox(NULL, _T("Worksheet not found."), _T("Error"), MB_OK | MB_ICONSTOP);
//		return FALSE;
//	}
//	
//
//	return TRUE;
//}


BOOL CXLAutomation::AddNewSheet(CString sheetName)
{
	if (m_pdispExcelApp == NULL || m_pdispWorkbook == NULL)
		return FALSE;

	VARIANTARG vargSheets, vargNewSheet;
	VariantInit(&vargSheets);
	VariantInit(&vargNewSheet);

	// Get the Worksheets collection
	if (!ExlInvoke(m_pdispWorkbook, L"Worksheets", &vargSheets, DISPATCH_PROPERTYGET, 0))
	{
		MessageBox(NULL, _T("Failed to get Worksheets collection."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	// Add a new sheet
	ClearAllArgs();
	if (!ExlInvoke(vargSheets.pdispVal, L"Add", &vargNewSheet, DISPATCH_METHOD, 0))
	{
		MessageBox(NULL, _T("Failed to add new sheet."), _T("Error"), MB_OK | MB_ICONERROR);
		VariantClear(&vargSheets);
		return FALSE;
	}

	// Set the sheet name
	ClearAllArgs();
	AddArgumentCString(NULL, 0, sheetName);
	if (!ExlInvoke(vargNewSheet.pdispVal, L"Name", NULL, DISPATCH_PROPERTYPUT, 0))
	{
		MessageBox(NULL, _T("Failed to set sheet name."), _T("Error"), MB_OK | MB_ICONERROR);
		VariantClear(&vargSheets);
		VariantClear(&vargNewSheet);
		return FALSE;
	}

	VariantClear(&vargSheets);
	VariantClear(&vargNewSheet);

	return TRUE;
}

/////////////////////////////////////////////////////////
//song
BOOL CXLAutomation::FindAndStoreWorksheet(IDispatch* pWorkbook, LPOLESTR sheetName, IDispatch** ppSheet) {
	VARIANTARG vargSheet;
	ClearAllArgs();
	AddArgumentOLEString(NULL, 0, sheetName);

	if (!ExlInvoke(pWorkbook, L"Worksheets", &vargSheet, DISPATCH_PROPERTYGET, DISP_FREEARGS))
		return FALSE;

	*ppSheet = vargSheet.pdispVal;
	return TRUE;
}


BOOL CXLAutomation::GetRange(SheetName sheet, int startRow, int startCol, int endRow, int endCol, VARIANTARG* pRange)
{
	if (m_pdispWorksheets[sheet] == NULL)
		return FALSE;

	VARIANTARG vargStartCell, vargEndCell;
	VariantInit(&vargStartCell);
	VariantInit(&vargEndCell);

	// 첫 번째 셀 객체 가져오기 column,row, 순서로 전달)!!!!!!
	ClearAllArgs();	
	AddArgumentInt2(NULL, 0, startCol); // Column
	AddArgumentInt2(NULL, 0, startRow); // Row
	if (!ExlInvoke(m_pdispWorksheets[sheet], L"Cells", &vargStartCell, DISPATCH_PROPERTYGET, DISP_FREEARGS))
		return FALSE;

	// 마지막 셀 객체 가져오기 column,row, 순서로 전달)!!!!!!
	ClearAllArgs();	
	AddArgumentInt2(NULL, 0, endCol);   // Column
	AddArgumentInt2(NULL, 0, endRow);   // Row
	if (!ExlInvoke(m_pdispWorksheets[sheet], L"Cells", &vargEndCell, DISPATCH_PROPERTYGET, DISP_FREEARGS))
	{
		VariantClear(&vargStartCell);
		return FALSE;
	}

	// Range 메서드 호출하여 범위 설정
	ClearAllArgs();
	AddArgumentDispatch(NULL, 0, vargStartCell.pdispVal);
	AddArgumentDispatch(NULL, 0, vargEndCell.pdispVal);
	BOOL result = ExlInvoke(m_pdispWorksheets[sheet], L"Range", pRange, DISPATCH_PROPERTYGET, DISP_FREEARGS);

	VariantClear(&vargStartCell);
	VariantClear(&vargEndCell);

	return result;
}


BOOL CXLAutomation::ReadRangeToArray(SheetName sheet, int startRow, int startCol, int endRow, int endCol, int* dataArray, int rows, int cols)
{
	VARIANTARG vargRng, vargData;

	// 범위를 설정하고 Excel에서 해당 범위의 IDispatch 포인터를 가져옵니다.
	if (!GetRange(sheet, startRow, startCol, endRow, endCol, &vargRng))
		return FALSE;

	// Excel 범위의 데이터를 가져옴
	if (!ExlInvoke(vargRng.pdispVal, L"Value", &vargData, DISPATCH_PROPERTYGET, 0))
	{
		VariantClear(&vargRng);
		return FALSE;
	}

	// SAFEARRAY 접근
	SAFEARRAY* pSafeArray = vargData.parray;
	VARIANT* pVarData = NULL;
	HRESULT hr = SafeArrayAccessData(pSafeArray, (void**)&pVarData);
	if (FAILED(hr))
	{
		VariantClear(&vargRng);
		VariantClear(&vargData);
		MessageBox(NULL, _T("Failed to access SafeArray data."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	// SAFEARRAY에서 데이터 추출
	LONG lRowLBound, lRowUBound, lColLBound, lColUBound;
	SafeArrayGetLBound(pSafeArray, 1, &lRowLBound);
	SafeArrayGetUBound(pSafeArray, 1, &lRowUBound);
	SafeArrayGetLBound(pSafeArray, 2, &lColLBound);
	SafeArrayGetUBound(pSafeArray, 2, &lColUBound);

	
	/*for (LONG r = lRowLBound; r <= lRowUBound; r++)
	{*/
		//for (LONG c = lColLBound; c <= lColUBound; c++)		
		ULONG lLoop = (lRowUBound- lRowLBound+1) * (lColUBound - lColLBound + 1);
		for (ULONG c = 0; c <= lLoop ; c++)
		{
//			LONG index[2] = { r, c };
			VARIANT* pVarCell = &pVarData[c];

			// 데이터 타입에 따라 처리
			if (pVarCell->vt == VT_I4) // Integer
			{
				*(dataArray + c) = pVarCell->lVal;
			}
			else if (pVarCell->vt == VT_R8) // Double, in case of float numbers
			{
				*(dataArray + c) = (int)pVarCell->dblVal;
			}
			else if (pVarCell->vt == VT_BSTR) // String, if needed to handle
			{
				// 문자열을 정수로 변환하거나 필요에 따라 처리할 수 있습니다.
			}
			else if (pVarCell->vt == VT_EMPTY) // Empty cell
			{
				*(dataArray + c) = 0; // Or any default value
			}
		}
	//}

	SafeArrayUnaccessData(pSafeArray);
	VariantClear(&vargRng);
	VariantClear(&vargData);
	return TRUE;
}

// 엑셀 셀 값 가져오기 (기존 Variant 리턴 함수)
BOOL CXLAutomation::GetCellValueVariant(SheetName sheet, int nRow, int nColumn, VARIANTARG* pValue)
{
	if (m_pdispWorksheets[sheet] == NULL)
		return false;

	VARIANTARG vargCell;
	VariantInit(&vargCell);

	ClearAllArgs();	
	AddArgumentDouble(NULL, 0, nColumn);	
	AddArgumentDouble(NULL, 0, nRow);
	if (!ExlInvoke(m_pdispWorksheets[sheet], L"Cells", &vargCell, DISPATCH_PROPERTYGET, DISP_FREEARGS))
		return false;

	if (!ExlInvoke(vargCell.pdispVal, L"Value", pValue, DISPATCH_PROPERTYGET, 0))
	{
		VariantClear(&vargCell);
		return false;
	}

	VariantClear(&vargCell);
	return true;
}

BOOL CXLAutomation::GetCellValueInt(SheetName sheet, int nRow, int nColumn, int* result)
{
	if (m_pdispWorksheets[sheet] == NULL)
		return FALSE;

	VARIANTARG vargRng, vargValue;
	VariantInit(&vargRng);
	VariantInit(&vargValue);

	// 해당 셀 범위 가져오기
	ClearAllArgs();		
	AddArgumentDouble(NULL, 0, nColumn); // Column
	AddArgumentDouble(NULL, 0, nRow);    // Row
	if (!ExlInvoke(m_pdispWorksheets[sheet], L"Cells", &vargRng, DISPATCH_PROPERTYGET, DISP_FREEARGS))
		return FALSE;

	// 셀 값 가져오기
	if (!ExlInvoke(vargRng.pdispVal, L"Value", &vargValue, DISPATCH_PROPERTYGET, 0))
	{
		VariantClear(&vargRng);
		return FALSE;
	}

	// VARIANT 타입에 따라 처리
	if (vargValue.vt == VT_I4) // Integer
	{
		*result = vargValue.lVal;
	}
	else if (vargValue.vt == VT_R8) // Double
	{
		*result = static_cast<int>(vargValue.dblVal); // Double을 Int로 변환
	}
	else if (vargValue.vt == VT_R4) // Float
	{
		*result = static_cast<int>(vargValue.fltVal); // Float을 Int로 변환
	}
	else
	{
		VariantClear(&vargRng);
		VariantClear(&vargValue);
		return FALSE; // 지원하지 않는 타입일 경우 FALSE 반환
	}

	VariantClear(&vargRng);
	VariantClear(&vargValue);
	return TRUE;
}

// Get double value from Worksheet.Cells(nColumn, nRow)
BOOL CXLAutomation::GetCellValueDouble(SheetName sheet, int nRow, int nColumn, double* result)
{
	if (m_pdispWorksheets[sheet] == NULL)
		return FALSE;

	VARIANTARG vargRng, vargValue;
	VariantInit(&vargRng);
	VariantInit(&vargValue);

	// 해당 셀 범위 가져오기
	ClearAllArgs();		
	AddArgumentDouble(NULL, 0, nColumn); // Column
	AddArgumentDouble(NULL, 0, nRow);    // Row
	if (!ExlInvoke(m_pdispWorksheets[sheet], L"Cells", &vargRng, DISPATCH_PROPERTYGET, DISP_FREEARGS))
		return FALSE;

	// 셀 값 가져오기
	if (!ExlInvoke(vargRng.pdispVal, L"Value", &vargValue, DISPATCH_PROPERTYGET, 0))
	{
		VariantClear(&vargRng);
		return FALSE;
	}

	// VARIANT 타입에 따라 처리
	if (vargValue.vt == VT_R8) // Double
	{
		*result = vargValue.dblVal;
	}
	else if (vargValue.vt == VT_R4) // Float
	{
		*result = static_cast<double>(vargValue.fltVal); // Float을 Double로 변환
	}
	else if (vargValue.vt == VT_I4) // Integer
	{
		*result = static_cast<double>(vargValue.lVal); // Integer를 Double로 변환
	}
	else
	{
		VariantClear(&vargRng);
		VariantClear(&vargValue);
		return FALSE; // 지원하지 않는 타입일 경우 FALSE 반환
	}

	VariantClear(&vargRng);
	VariantClear(&vargValue);
	return TRUE;
}


// Get CString value from Worksheet.Cells(nColumn, nRow)
BOOL CXLAutomation::GetCellValueCString(SheetName sheet, int nRow, int nColumn, CString* pValue)
{
	if (NULL == m_pdispWorksheets[sheet] || pValue == nullptr)
		return FALSE;

	VARIANTARG vargRng, vargValue;
	VariantInit(&vargValue);

	ClearAllArgs();		
	AddArgumentDouble(NULL, 0, nColumn);
	AddArgumentDouble(NULL, 0, nRow);
	if (!ExlInvoke(m_pdispWorksheets[sheet], L"Cells", &vargRng, DISPATCH_PROPERTYGET, DISP_FREEARGS))
		return FALSE;

	if (!ExlInvoke(vargRng.pdispVal, L"Value", &vargValue, DISPATCH_PROPERTYGET, 0))
		return FALSE;

	if (vargValue.vt == VT_BSTR) // Check if the value is a string
	{
		*pValue = vargValue.bstrVal;
	}
	else
	{
		MessageBox(NULL, _T("The cell value is not a string."), _T("Type Error"), MB_OK | MB_ICONERROR);
		VariantClear(&vargRng);
		VariantClear(&vargValue);
		return FALSE;
	}

	VariantClear(&vargRng);
	VariantClear(&vargValue);
	return TRUE;
}


// SetCellValue for integer
BOOL CXLAutomation::SetCellValueInt(SheetName sheet, int nRow, int nColumn, int value)
{
	if (m_pdispWorksheets[sheet] == NULL)
		return FALSE;

	VARIANTARG vargRng;
	VariantInit(&vargRng);

	// 해당 셀 범위 가져오기
	ClearAllArgs();		
	AddArgumentDouble(NULL, 0, nColumn); // Column
	AddArgumentDouble(NULL, 0, nRow);    // Row
	if (!ExlInvoke(m_pdispWorksheets[sheet], L"Cells", &vargRng, DISPATCH_PROPERTYGET, DISP_FREEARGS))
		return FALSE;

	// 셀 값 설정 (정수)
	ClearAllArgs();
	AddArgumentInt2(NULL, 0, value); // 정수 값을 설정
	if (!ExlInvoke(vargRng.pdispVal, L"Value", NULL, DISPATCH_PROPERTYPUT, 0))
	{
		VariantClear(&vargRng);
		return FALSE;
	}

	VariantClear(&vargRng);
	return TRUE;
}

// SetCellValue for CString
BOOL CXLAutomation::SetCellValueCString(SheetName sheet, int nRow, int nColumn, CString value)
{
	if (m_pdispWorksheets[sheet] == NULL)
		return FALSE;

	VARIANTARG vargRng;
	VariantInit(&vargRng);

	// 해당 셀 범위 가져오기
	ClearAllArgs();		
	AddArgumentDouble(NULL, 0, nColumn); // Column
	AddArgumentDouble(NULL, 0, nRow);    // Row
	if (!ExlInvoke(m_pdispWorksheets[sheet], L"Cells", &vargRng, DISPATCH_PROPERTYGET, DISP_FREEARGS))
		return FALSE;

	// 셀 값 설정 (문자열)
	ClearAllArgs();
	AddArgumentCString(NULL, 0, value); // CString 값을 설정
	if (!ExlInvoke(vargRng.pdispVal, L"Value", NULL, DISPATCH_PROPERTYPUT, 0))
	{
		VariantClear(&vargRng);
		return FALSE;
	}

	VariantClear(&vargRng);
	return TRUE;
}

// SetCellValue for double
BOOL CXLAutomation::SetCellValueDouble(SheetName sheet, int nRow, int nColumn, double value)
{
	if (m_pdispWorksheets[sheet] == NULL)
		return FALSE;

	VARIANTARG vargRng;
	VariantInit(&vargRng);

	// 해당 셀 범위 가져오기
	ClearAllArgs();
	AddArgumentDouble(NULL, 0, nColumn); // Column	
	AddArgumentDouble(NULL, 0, nRow);    // Row	
	if (!ExlInvoke(m_pdispWorksheets[sheet], L"Cells", &vargRng, DISPATCH_PROPERTYGET, DISP_FREEARGS))
		return FALSE;

	// 셀 값 설정 (더블)
	ClearAllArgs();
	AddArgumentDouble(NULL, 0, value); // 더블 값을 설정
	if (!ExlInvoke(vargRng.pdispVal, L"Value", NULL, DISPATCH_PROPERTYPUT, 0))
	{
		VariantClear(&vargRng);
		return FALSE;
	}

	VariantClear(&vargRng);
	return TRUE;
}

// ReadRangeToIntArray: Excel 범위 데이터를 int 배열로 읽어오기
BOOL CXLAutomation::ReadRangeToIntArray(SheetName sheet, int startRow, int startCol, int* dataArray, int rows, int cols) {
	if (m_pdispWorksheets[sheet] == NULL)
		return FALSE;

	VARIANTARG vargRng, vargData;
	VariantInit(&vargRng);
	VariantInit(&vargData);

	// Get the range from Excel
	if (!GetRange(sheet, startRow, startCol, rows, startCol + cols - 1, &vargRng)) {
		MessageBox(NULL, _T("Failed to get Excel range."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	// Read the data from the range
	if (!ExlInvoke(vargRng.pdispVal, L"Value", &vargData, DISPATCH_PROPERTYGET, 0)) {
		VariantClear(&vargRng);
		return FALSE;
	}

	// Access the SAFEARRAY data
	SAFEARRAY* pSafeArray = vargData.parray;
	VARIANT* pVarData = NULL;
	HRESULT hr = SafeArrayAccessData(pSafeArray, (void**)&pVarData);
	if (FAILED(hr)) {
		VariantClear(&vargRng);
		VariantClear(&vargData);
		MessageBox(NULL, _T("Failed to access SafeArray data."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	// Extract data from SAFEARRAY
	LONG lRowLBound, lRowUBound, lColLBound, lColUBound;
	SafeArrayGetLBound(pSafeArray, 1, &lRowLBound);
	SafeArrayGetUBound(pSafeArray, 1, &lRowUBound);
	SafeArrayGetLBound(pSafeArray, 2, &lColLBound);
	SafeArrayGetUBound(pSafeArray, 2, &lColUBound);

	for (LONG r = 0; r < rows; r++) {
		for (LONG c = 0; c < cols; c++) {
			LONG index[2] = { r + lRowLBound, c + lColLBound }; // Corrected indices for column-major order

			VARIANT varCell;
			VariantInit(&varCell);
			HRESULT hr = SafeArrayGetElement(pSafeArray, index, &varCell);

			if (FAILED(hr)) {
				SafeArrayUnaccessData(pSafeArray);
				VariantClear(&vargRng);
				VariantClear(&vargData);
				MessageBox(NULL, _T("Failed to get element from SafeArray."), _T("Error"), MB_OK | MB_ICONERROR);
				return FALSE;
			}

			if (varCell.vt == VT_I4) {
				dataArray[r * cols + c] = varCell.lVal;
			}
			else if (varCell.vt == VT_R8) {
				dataArray[r * cols + c] = static_cast<int>(varCell.dblVal);
			}
			else if (varCell.vt == VT_R4) {
				dataArray[r * cols + c] = static_cast<int>(varCell.fltVal);
			}
			else if (varCell.vt == VT_EMPTY) {
				dataArray[r * cols + c] = 0;
			}
			else {
				MessageBox(NULL, _T("Unexpected data type in Excel range."), _T("Error"), MB_OK | MB_ICONERROR);
				SafeArrayUnaccessData(pSafeArray);
				VariantClear(&varCell);
				VariantClear(&vargRng);
				VariantClear(&vargData);
				return FALSE;
			}

			VariantClear(&varCell);
		}
	}

	SafeArrayUnaccessData(pSafeArray);
	VariantClear(&vargRng);
	VariantClear(&vargData);
	return TRUE;
}


// ReadRangeToCStringArray: Excel 범위 데이터를 CString 배열로 읽어오기
BOOL CXLAutomation::ReadRangeToCStringArray(SheetName sheet, int startRow, int startCol, CString* dataArray, int rows, int cols) {
	if (m_pdispWorksheets[sheet] == NULL)
		return FALSE;

	VARIANTARG vargRng, vargData;
	VariantInit(&vargRng);
	VariantInit(&vargData);

	// Get the range from Excel
	if (!GetRange(sheet, startRow, startCol, rows , cols, &vargRng)) {
		MessageBox(NULL, _T("Failed to get Excel range."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	// Read the data from the range
	if (!ExlInvoke(vargRng.pdispVal, L"Value", &vargData, DISPATCH_PROPERTYGET, 0)) {
		VariantClear(&vargRng);
		return FALSE;
	}

	// Access the SAFEARRAY data
	SAFEARRAY* pSafeArray = vargData.parray;
	VARIANT* pVarData = NULL;
	HRESULT hr = SafeArrayAccessData(pSafeArray, (void**)&pVarData);
	if (FAILED(hr)) {
		VariantClear(&vargRng);
		VariantClear(&vargData);
		MessageBox(NULL, _T("Failed to access SafeArray data."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	// Extract data from SAFEARRAY using 1-based indexing
	LONG lRowLBound, lRowUBound, lColLBound, lColUBound;
	SafeArrayGetLBound(pSafeArray, 1, &lRowLBound);
	SafeArrayGetUBound(pSafeArray, 1, &lRowUBound);
	SafeArrayGetLBound(pSafeArray, 2, &lColLBound);
	SafeArrayGetUBound(pSafeArray, 2, &lColUBound);

	for (LONG r = 0; r < rows; r++) {
		for (LONG c = 0; c < cols; c++) {
			LONG index[2] = { r + lRowLBound, c + lColLBound }; // Corrected indices for column-major order

			VARIANT varCell;
			VariantInit(&varCell);
			HRESULT hr = SafeArrayGetElement(pSafeArray, index, &varCell);

			if (FAILED(hr)) {
				SafeArrayUnaccessData(pSafeArray);
				VariantClear(&vargRng);
				VariantClear(&vargData);
				MessageBox(NULL, _T("Failed to get element from SafeArray."), _T("Error"), MB_OK | MB_ICONERROR);
				return FALSE;
			}

			// Check if the VARIANT is a string or empty, otherwise error
			if (varCell.vt == VT_BSTR) {
				dataArray[r * cols + c] = CString(varCell.bstrVal);
			}
			else if (varCell.vt == VT_EMPTY) {
				dataArray[r * cols + c] = _T("");
			}
			else {
				MessageBox(NULL, _T("Non-string data type found in Excel range."), _T("Error"), MB_OK | MB_ICONERROR);
				SafeArrayUnaccessData(pSafeArray);
				VariantClear(&varCell);
				VariantClear(&vargRng);
				VariantClear(&vargData);
				return FALSE;
			}

			VariantClear(&varCell); // Clear the VARIANT after processing
		}
	}

	SafeArrayUnaccessData(pSafeArray);
	VariantClear(&vargRng);
	VariantClear(&vargData);
	return TRUE;
}

BOOL CXLAutomation::WriteArrayToRangeInt(SheetName sheet, int startRow, int startCol, int* dataArray, int rows, int cols) {
	if (m_pdispWorksheets[sheet] == NULL)
		return FALSE;

	VARIANTARG vargRng;
	VariantInit(&vargRng);

	// Set the Excel range using the GetRange function
	if (!GetRange(sheet, startRow, startCol, rows, cols, &vargRng)) {
		MessageBox(NULL, _T("Failed to get Excel range."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	// Create a SAFEARRAY for the data
	SAFEARRAYBOUND sab[2];
	sab[0].lLbound = 1;
	sab[0].cElements = rows;
	sab[1].lLbound = 1;
	sab[1].cElements = cols;

	SAFEARRAY* pSafeArray = SafeArrayCreate(VT_VARIANT, 2, sab);
	if (pSafeArray == NULL) {
		MessageBox(NULL, _T("Failed to create SAFEARRAY."), _T("Error"), MB_OK | MB_ICONERROR);
		VariantClear(&vargRng);
		return FALSE;
	}

	// Fill the SAFEARRAY with data
	for (long r = 1; r <= rows; ++r) {
		for (long c = 1; c <= cols; ++c) {
			VARIANT vtData;
			VariantInit(&vtData);
			vtData.vt = VT_I4; // Integer format
			vtData.lVal = *(dataArray + (r - 1) * cols + (c - 1)); // Note: Indexing corrected to match 1-based indexing of SAFEARRAY

			long indices[2] = { r, c }; // SAFEARRAY is 1-based
			SafeArrayPutElement(pSafeArray, indices, &vtData);
		}
	}

	// Wrap the SAFEARRAY in a VARIANT
	VARIANT vtArray;
	VariantInit(&vtArray);
	vtArray.vt = VT_ARRAY | VT_VARIANT;
	vtArray.parray = pSafeArray;
	
	// Write the SAFEARRAY data to the range
	HRESULT hr = AutoWrap(DISPATCH_PROPERTYPUT, NULL, vargRng.pdispVal, L"Value", 1, vtArray);

	// Clean up
	VariantClear(&vargRng);
	SafeArrayDestroy(pSafeArray);

	// Check for errors
	if (FAILED(hr)) {
		MessageBox(NULL, _T("Failed to write data to Excel range."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	return TRUE;
}

BOOL CXLAutomation::WriteArrayToRangeCString(SheetName sheet, int startRow, int startCol, CString* dataArray, int rows, int cols) {
	if (m_pdispWorksheets[sheet] == NULL)
		return FALSE;

	VARIANTARG vargRng;
	VariantInit(&vargRng);

	// Set the Excel range using the GetRange function
	//if (!GetRange(sheet, startRow, startCol, rows, cols, &vargRng)) {
	if (!GetRange(sheet, startRow, startCol, startRow + rows - 1, startCol + cols - 1, &vargRng)){
		MessageBox(NULL, _T("Failed to get Excel range."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	// Create a SAFEARRAY for the data
	SAFEARRAYBOUND sab[2];
	sab[0].lLbound = 1;
	sab[0].cElements = rows;
	sab[1].lLbound = 1;
	sab[1].cElements = cols;

	SAFEARRAY* pSafeArray = SafeArrayCreate(VT_VARIANT, 2, sab);
	if (pSafeArray == NULL) {
		MessageBox(NULL, _T("Failed to create SAFEARRAY."), _T("Error"), MB_OK | MB_ICONERROR);
		VariantClear(&vargRng);
		return FALSE;
	}

	// Fill the SAFEARRAY with data
	// Fill the SAFEARRAY with data
	for (long r = 1; r <= rows; ++r) {
		for (long c = 1; c <= cols; ++c) {
			VARIANT vtData;
			VariantInit(&vtData);

			// Check if CString is empty
			if (dataArray[(r - 1) * cols + (c - 1)].IsEmpty()) {
				// Set the VARIANT type to VT_EMPTY if the CString is empty
				vtData.vt = VT_EMPTY;
			}
			else {
				// Otherwise, use the CString value
				vtData.vt = VT_BSTR; // String format
				vtData.bstrVal = SysAllocString(dataArray[(r - 1) * cols + (c - 1)]); // Convert CString to BSTR

				if (vtData.bstrVal == NULL) {
					MessageBox(NULL, _T("Failed to allocate memory for BSTR."), _T("Error"), MB_OK | MB_ICONERROR);
					SafeArrayDestroy(pSafeArray);
					VariantClear(&vargRng);
					return FALSE;
				}
			}

			long indices[2] = { r, c }; // SAFEARRAY is 1-based
			SafeArrayPutElement(pSafeArray, indices, &vtData);
			VariantClear(&vtData); // Clear VARIANT after use
		}
	}


	// Wrap the SAFEARRAY in a VARIANT
	VARIANT vtArray;
	VariantInit(&vtArray);
	vtArray.vt = VT_ARRAY | VT_VARIANT;
	vtArray.parray = pSafeArray;

	// Write the SAFEARRAY data to the range
	HRESULT hr = AutoWrap(DISPATCH_PROPERTYPUT, NULL, vargRng.pdispVal, L"Value", 1, vtArray);

	// Clean up
	VariantClear(&vargRng);
	SafeArrayDestroy(pSafeArray);

	// Check for errors
	if (FAILED(hr)) {
		MessageBox(NULL, _T("Failed to write data to Excel range."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	return TRUE;
}


BOOL CXLAutomation::WriteArrayToRangeVariant(SheetName sheet, int startRow, int startCol, VARIANT* dataArray, int rows, int cols) {
	if (m_pdispWorksheets[sheet] == NULL)
		return FALSE;

	VARIANTARG vargRng;
	VariantInit(&vargRng);

	// Set the Excel range using the GetRange function
	if (!GetRange(sheet, startRow, startCol, rows, cols , &vargRng)) {
		MessageBox(NULL, _T("Failed to get Excel range."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	// Create a SAFEARRAY for the data
	SAFEARRAYBOUND sab[2];
	sab[0].lLbound = 1;
	sab[0].cElements = rows;
	sab[1].lLbound = 1;
	sab[1].cElements = cols;

	SAFEARRAY* pSafeArray = SafeArrayCreate(VT_VARIANT, 2, sab);
	if (pSafeArray == NULL) {
		MessageBox(NULL, _T("Failed to create SAFEARRAY."), _T("Error"), MB_OK | MB_ICONERROR);
		VariantClear(&vargRng);
		return FALSE;
	}

	// Fill the SAFEARRAY with data
	for (long r = 1; r <= rows; ++r) {
		for (long c = 1; c <= cols; ++c) {
			VARIANT vtData;
			VariantInit(&vtData);

			// Handle each VARIANT type
			switch (dataArray[(r - 1) * cols + (c - 1)].vt) {
			case VT_INT:
				vtData.vt = VT_INT;
				vtData.intVal = dataArray[(r - 1) * cols + (c - 1)].intVal;
				break;
			case VT_I4:
				vtData.vt = VT_I4;
				vtData.lVal = dataArray[(r - 1) * cols + (c - 1)].lVal;
				break;
			case VT_R8:
				vtData.vt = VT_R8;
				vtData.dblVal = dataArray[(r - 1) * cols + (c - 1)].dblVal;
				break;
			case VT_BOOL:
				vtData.vt = VT_BOOL;
				vtData.boolVal = dataArray[(r - 1) * cols + (c - 1)].boolVal;
				break;
			case VT_BSTR:
				vtData.vt = VT_BSTR;
				vtData.bstrVal = SysAllocString(dataArray[(r - 1) * cols + (c - 1)].bstrVal);
				if (vtData.bstrVal == NULL) {
					MessageBox(NULL, _T("Failed to allocate memory for BSTR."), _T("Error"), MB_OK | MB_ICONERROR);
					SafeArrayDestroy(pSafeArray);
					VariantClear(&vargRng);
					return FALSE;
				}
				break;
			case VT_UI1:
				vtData.vt = VT_UI1;
				vtData.bVal = dataArray[(r - 1) * cols + (c - 1)].bVal;
				break;
			case VT_I2:
				vtData.vt = VT_I2;
				vtData.iVal = dataArray[(r - 1) * cols + (c - 1)].iVal;
				break;
			case VT_UI2:
				vtData.vt = VT_UI2;
				vtData.uiVal = dataArray[(r - 1) * cols + (c - 1)].uiVal;
				break;
			case VT_UI4:
				vtData.vt = VT_UI4;
				vtData.ulVal = dataArray[(r - 1) * cols + (c - 1)].ulVal;
				break;
			case VT_I8:
				vtData.vt = VT_I8;
				vtData.llVal = dataArray[(r - 1) * cols + (c - 1)].llVal;
				break;
			case VT_UI8:
				vtData.vt = VT_UI8;
				vtData.ullVal = dataArray[(r - 1) * cols + (c - 1)].ullVal;
				break;
			case VT_R4:
				vtData.vt = VT_R4;
				vtData.fltVal = dataArray[(r - 1) * cols + (c - 1)].fltVal;
				break;
			case VT_DATE:
				vtData.vt = VT_DATE;
				vtData.date = dataArray[(r - 1) * cols + (c - 1)].date;
				break;
			case VT_CY:
				vtData.vt = VT_CY;
				vtData.cyVal = dataArray[(r - 1) * cols + (c - 1)].cyVal;
				break;
			case VT_EMPTY:
				vtData.vt = VT_EMPTY;
				break;
				// Add additional cases for other supported types if needed
			default:
				MessageBox(NULL, _T("Unsupported VARIANT type in data array."), _T("Error"), MB_OK | MB_ICONERROR);
				SafeArrayDestroy(pSafeArray);
				VariantClear(&vargRng);
				return FALSE;
			}

			long indices[2] = { r, c }; // SAFEARRAY is 1-based
			SafeArrayPutElement(pSafeArray, indices, &vtData);
			VariantClear(&vtData); // Clear VARIANT after use
		}
	}

	// Wrap the SAFEARRAY in a VARIANT
	VARIANT vtArray;
	VariantInit(&vtArray);
	vtArray.vt = VT_ARRAY | VT_VARIANT;
	vtArray.parray = pSafeArray;

	// Write the SAFEARRAY data to the range
	HRESULT hr = AutoWrap(DISPATCH_PROPERTYPUT, NULL, vargRng.pdispVal, L"Value", 1, vtArray);

	// Clean up
	VariantClear(&vargRng);
	SafeArrayDestroy(pSafeArray);

	// Check for errors
	if (FAILED(hr)) {
		MessageBox(NULL, _T("Failed to write data to Excel range."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	return TRUE;
}


// Define the AutoWrap function
HRESULT CXLAutomation::AutoWrap(int autoType, VARIANT* pvResult, IDispatch* pDisp, LPOLESTR ptName, int cArgs...) {
	// 시작 변수-인수 목록...
	va_list marker;
	va_start(marker, cArgs);

	if (!pDisp) {
		MessageBox(NULL, _T("NULL IDispatch passed to AutoWrap()"), _T("Error"), MB_OK | MB_ICONERROR);
		return E_INVALIDARG;
	}

	// 사용된 변수들...
	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	DISPID dispID;
	HRESULT hr;
	WCHAR buf[200];
	char szName[200];

	// ANSI로 변환
	WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);

	// 전달된 이름에 대한 DISPID 가져오기...
	hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
	if (FAILED(hr)) {
		wsprintf(buf, _T("IDispatch::GetIDsOfNames(\"%s\") failed w/err 0x%08lx"), ptName, hr);
		MessageBox(NULL, buf, _T("AutoWrap()"), MB_OK | MB_ICONERROR);
		return hr;
	}

	// 인수에 대한 메모리 할당...
	VARIANT* pArgs = new VARIANT[cArgs + 1];
	// 인수 추출...
	for (int i = 0; i < cArgs; i++) {
		pArgs[i] = va_arg(marker, VARIANT);
	}

	// DISPPARAMS 생성
	dp.cArgs = cArgs;
	dp.rgvarg = pArgs;

	// property-put의 특별한 경우 처리!
	if (autoType & DISPATCH_PROPERTYPUT) {
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispidNamed;
	}

	// 호출 수행!
	hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
	if (FAILED(hr)) {
		wsprintf(buf, _T("IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx"), ptName, dispID, hr);
		MessageBox(NULL, buf, _T("AutoWrap()"), MB_OK | MB_ICONERROR);
		return hr;
	}

	// 가변 인수 섹션 종료...
	va_end(marker);

	delete[] pArgs;

	return hr;
}


/*
borderStyle: 엑셀의 xlContinuous, xlDash, xlDot 등과 같은 테두리 스타일 값을 전달합니다.
borderWeight: 엑셀의 xlThin, xlMedium, xlThick와 같은 테두리 두께 값을 전달합니다.
borderColor: RGB 값 또는 엑셀의 표준 색상 인덱스를 사용하여 테두리 색상을 설정합니다.

***엑셀 테두리 값 상수**
테두리 스타일 (LineStyle)
	xlContinuous = 1 (연속선)
	xlDash = -4115 (짧은 대시)
	xlDashDot = 4 (대시와 점)
	xlDashDotDot = 5 (대시와 두 점)
	xlDot = -4118 (점선)
	xlDouble = -4119 (이중선)
	xlLineStyleNone = -4142 (없음)
	xlSlantDashDot = 13 (경사 대시 점)

테두리 두께 (Weight)
	xlHairline = 1 (얇은 선)
	xlThin = 2 (얇은 테두리)
	xlMedium = -4138 (중간 두께)
	xlThick = 4 (두꺼운 테두리)
	
사용예제
	xlAutomation.SetRangeBorder(PROJECT, 1, 1, 5, 5, 1, 2, RGB(255, 0, 0));
	PROJECT 시트의 (1,1)에서 (5,5) 범위에 연속선 스타일(1), 얇은 두께(2), 빨간색(RGB(255, 0, 0)) 테두리를 설정
*/
BOOL CXLAutomation::SetRangeBorder(SheetName sheet, int startRow, int startCol, int endRow, int endCol, int borderStyle, int borderWeight, int borderColor)
{
	if (m_pdispWorksheets[sheet] == NULL)
		return FALSE;

	VARIANTARG vargRng;
	VariantInit(&vargRng);

	// Get the Excel range
	if (!GetRange(sheet, startRow, startCol, endRow, endCol, &vargRng))
	{
		MessageBox(NULL, _T("Failed to get Excel range."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	// Set borders for all edges (xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
	int borders[] = { 7, 8, 9, 10, 11, 12 }; // xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal

	for (int i = 0; i < sizeof(borders) / sizeof(borders[0]); i++)
	{
		VARIANTARG vargBorder;
		VariantInit(&vargBorder);

		// Access the Borders property of the Range
		ClearAllArgs();
		AddArgumentInt2(NULL, 0, borders[i]);
		if (!ExlInvoke(vargRng.pdispVal, L"Borders", &vargBorder, DISPATCH_PROPERTYGET, DISP_FREEARGS))
		{
			VariantClear(&vargRng);
			return FALSE;
		}

		// Set Border Style
		ClearAllArgs();
		AddArgumentInt2(NULL, 0, borderStyle);
		if (!ExlInvoke(vargBorder.pdispVal, L"LineStyle", NULL, DISPATCH_PROPERTYPUT, 0))
		{
			VariantClear(&vargRng);
			VariantClear(&vargBorder);
			return FALSE;
		}

		// Set Border Weight
		ClearAllArgs();
		AddArgumentInt2(NULL, 0, borderWeight);
		if (!ExlInvoke(vargBorder.pdispVal, L"Weight", NULL, DISPATCH_PROPERTYPUT, 0))
		{
			VariantClear(&vargRng);
			VariantClear(&vargBorder);
			return FALSE;
		}

		// Set Border Color
		ClearAllArgs();
		AddArgumentInt2(NULL, 0, borderColor);
		if (!ExlInvoke(vargBorder.pdispVal, L"Color", NULL, DISPATCH_PROPERTYPUT, 0))
		{
			VariantClear(&vargRng);
			VariantClear(&vargBorder);
			return FALSE;
		}

		VariantClear(&vargBorder);
	}

	VariantClear(&vargRng);
	return TRUE;
}

BOOL CXLAutomation::SetRangeBorderAround(SheetName sheet, int startRow, int startCol, int endRow, int endCol, int borderStyle, int borderWeight, int borderColor)
{
	if (m_pdispWorksheets[sheet] == NULL)
		return FALSE;

	VARIANTARG vargRng;
	VariantInit(&vargRng);

	// Get the Excel range
	if (!GetRange(sheet, startRow, startCol, endRow, endCol, &vargRng))
	{
		MessageBox(NULL, _T("Failed to get Excel range."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	// Define the border edges for around (xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
	int borders[] = { 7, 8, 9, 10 }; // xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight

	for (int i = 0; i < sizeof(borders) / sizeof(borders[0]); i++)
	{
		VARIANTARG vargBorder;
		VariantInit(&vargBorder);

		// Access the Borders property of the Range
		ClearAllArgs();
		AddArgumentInt2(NULL, 0, borders[i]);
		if (!ExlInvoke(vargRng.pdispVal, L"Borders", &vargBorder, DISPATCH_PROPERTYGET, DISP_FREEARGS))
		{
			VariantClear(&vargRng);
			return FALSE;
		}

		// Set Border Style
		ClearAllArgs();
		AddArgumentInt2(NULL, 0, borderStyle);
		if (!ExlInvoke(vargBorder.pdispVal, L"LineStyle", NULL, DISPATCH_PROPERTYPUT, 0))
		{
			VariantClear(&vargRng);
			VariantClear(&vargBorder);
			return FALSE;
		}

		// Set Border Weight
		ClearAllArgs();
		AddArgumentInt2(NULL, 0, borderWeight);
		if (!ExlInvoke(vargBorder.pdispVal, L"Weight", NULL, DISPATCH_PROPERTYPUT, 0))
		{
			VariantClear(&vargRng);
			VariantClear(&vargBorder);
			return FALSE;
		}

		// Set Border Color
		ClearAllArgs();
		AddArgumentInt2(NULL, 0, borderColor);
		if (!ExlInvoke(vargBorder.pdispVal, L"Color", NULL, DISPATCH_PROPERTYPUT, 0))
		{
			VariantClear(&vargRng);
			VariantClear(&vargBorder);
			return FALSE;
		}

		VariantClear(&vargBorder);
	}

	VariantClear(&vargRng);
	return TRUE;
}

BOOL CXLAutomation::ReadExRangeConvertInt(SheetName sheet, int startRow, int startCol, int* dataArray, int rows, int cols) {
	if (m_pdispWorksheets[sheet] == NULL)
		return FALSE;

	VARIANTARG vargRng, vargData;
	VariantInit(&vargRng);
	VariantInit(&vargData);

	// Get the range from Excel
	if (!GetRange(sheet, startRow, startCol, rows, cols, &vargRng)) {
		MessageBox(NULL, _T("Failed to get Excel range."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	// Read the data from the range
	if (!ExlInvoke(vargRng.pdispVal, L"Value", &vargData, DISPATCH_PROPERTYGET, 0)) {
		VariantClear(&vargRng);
		return FALSE;
	}

	// Access the SAFEARRAY data
	SAFEARRAY* pSafeArray = vargData.parray;
	VARIANT* pVarData = NULL;
	HRESULT hr = SafeArrayAccessData(pSafeArray, (void**)&pVarData);
	if (FAILED(hr)) {
		VariantClear(&vargRng);
		VariantClear(&vargData);
		MessageBox(NULL, _T("Failed to access SafeArray data."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	// Extract data from SAFEARRAY
	LONG lRowLBound, lRowUBound, lColLBound, lColUBound;
	SafeArrayGetLBound(pSafeArray, 1, &lRowLBound);
	SafeArrayGetUBound(pSafeArray, 1, &lRowUBound);
	SafeArrayGetLBound(pSafeArray, 2, &lColLBound);
	SafeArrayGetUBound(pSafeArray, 2, &lColUBound);

	for (LONG r = 0; r < rows; r++) {
		for (LONG c = 0; c < cols; c++) {
			LONG index[2] = { r + lRowLBound, c + lColLBound }; // Corrected indices for column-major order

			VARIANT varCell;
			VariantInit(&varCell);
			HRESULT hr = SafeArrayGetElement(pSafeArray, index, &varCell);

			if (FAILED(hr)) {
				SafeArrayUnaccessData(pSafeArray);
				VariantClear(&vargRng);
				VariantClear(&vargData);
				MessageBox(NULL, _T("Failed to get element from SafeArray."), _T("Error"), MB_OK | MB_ICONERROR);
				return FALSE;
			}

			// Convert to int or set to 0 if not a number
			switch (varCell.vt) {
			case VT_I4:
				dataArray[r * cols + c] = varCell.lVal;
				break;
			case VT_I2:
				dataArray[r * cols + c] = static_cast<int>(varCell.iVal);
				break;
			case VT_INT:
				dataArray[r * cols + c] = varCell.intVal;
				break;
			case VT_UI1:
				dataArray[r * cols + c] = static_cast<int>(varCell.bVal);
				break;
			case VT_UI2:
				dataArray[r * cols + c] = static_cast<int>(varCell.uiVal);
				break;
			case VT_UI4:
				dataArray[r * cols + c] = static_cast<int>(varCell.ulVal);
				break;
			case VT_I8:
				dataArray[r * cols + c] = static_cast<int>(varCell.llVal);
				break;
			case VT_UI8:
				dataArray[r * cols + c] = static_cast<int>(varCell.ullVal);
				break;
			case VT_R4:
				dataArray[r * cols + c] = static_cast<int>(varCell.fltVal);
				break;
			case VT_R8:
				dataArray[r * cols + c] = static_cast<int>(varCell.dblVal);
				break;
			default:
				// Set to 0 for non-numeric types or unsupported types
				dataArray[r * cols + c] = 0;
				break;
			}

			VariantClear(&varCell);
		}
	}

	SafeArrayUnaccessData(pSafeArray);
	VariantClear(&vargRng);
	VariantClear(&vargData);
	return TRUE;
}

BOOL CXLAutomation::SaveAndCloseExcelFile(CString szFileName) {
	if (m_pdispWorkbook == NULL) {
		MessageBox(NULL, _T("No open workbook to save."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	// Save the workbook without asking for SaveAs
	ClearAllArgs();
	if (!ExlInvoke(m_pdispWorkbook, L"Save", NULL, DISPATCH_METHOD, 0)) {
		MessageBox(NULL, _T("Failed to save the workbook."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	// Close the workbook
	ClearAllArgs();
	AddArgumentBool(L"SaveChanges", 0, FALSE);  // Save changes before closing
	if (!ExlInvoke(m_pdispWorkbook, L"Close", NULL, DISPATCH_METHOD, DISP_FREEARGS)) {
		MessageBox(NULL, _T("Failed to close the workbook."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	m_pdispWorkbook->Release();
	m_pdispWorkbook = NULL;

	// Quit Excel Application
	if (m_pdispExcelApp != NULL) {
		ClearAllArgs();
		if (!ExlInvoke(m_pdispExcelApp, L"Quit", NULL, DISPATCH_METHOD, 0)) {
			MessageBox(NULL, _T("Failed to quit Excel."), _T("Error"), MB_OK | MB_ICONERROR);
			return FALSE;
		}
		m_pdispExcelApp->Release();
		m_pdispExcelApp = NULL;
	}

	return TRUE;
}


//
//BOOL CXLAutomation::SaveAndCloseExcelFile(CString szFileName) {
//	if (m_pdispWorkbook == NULL) {
//		MessageBox(NULL, _T("No open workbook to save."), _T("Error"), MB_OK | MB_ICONERROR);
//		return FALSE;
//	}
//
//	// Save the workbook
//	VARIANTARG vargFilename;
//	VariantInit(&vargFilename);
//	vargFilename.vt = VT_BSTR;
//	vargFilename.bstrVal = szFileName.AllocSysString();
//
//	ClearAllArgs();
//	AddArgumentOLEString(L"Filename", 0, vargFilename.bstrVal);
//	if (!ExlInvoke(m_pdispWorkbook, L"SaveAs", NULL, DISPATCH_METHOD, DISP_FREEARGS)) {
//		VariantClear(&vargFilename);
//		MessageBox(NULL, _T("Failed to save the workbook."), _T("Error"), MB_OK | MB_ICONERROR);
//		return FALSE;
//	}
//
//	VariantClear(&vargFilename);
//
//	// Close the workbook
//	ClearAllArgs();
//	AddArgumentBool(L"SaveChanges", 0, TRUE);  // Save changes before closing
//	if (!ExlInvoke(m_pdispWorkbook, L"Close", NULL, DISPATCH_METHOD, DISP_FREEARGS)) {
//		MessageBox(NULL, _T("Failed to close the workbook."), _T("Error"), MB_OK | MB_ICONERROR);
//		return FALSE;
//	}
//
//	m_pdispWorkbook->Release();
//	m_pdispWorkbook = NULL;
//
//	// Quit Excel Application
//	if (m_pdispExcelApp != NULL) {
//		ClearAllArgs();
//		if (!ExlInvoke(m_pdispExcelApp, L"Quit", NULL, DISPATCH_METHOD, 0)) {
//			MessageBox(NULL, _T("Failed to quit Excel."), _T("Error"), MB_OK | MB_ICONERROR);
//			return FALSE;
//		}
//		m_pdispExcelApp->Release();
//		m_pdispExcelApp = NULL;
//	}
//
//	return TRUE;
//}
