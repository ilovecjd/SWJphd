// XLAutomation.cpp: implementation of the CXLAutomation class.
//This is C++ modification of the AutoXL C-sample from 
//Microsoft Excel97 Developer Kit, Microsoft Press 1997 
//////////////////////////////////////////////////////////////////////

#include "pch.h"
//#include "XLAutomationTester.h"
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
	
	InitOLE();
	StartExcel();
	SetExcelVisible(TRUE);	
}

CXLAutomation::CXLAutomation(BOOL bVisible)
{
	m_pdispExcelApp = NULL;
	m_pdispWorkbook = NULL;
	
	InitOLE();
	StartExcel();
	SetExcelVisible(bVisible);
	
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
BOOL CXLAutomation::SetCellsValueToString(SheetName sheet, double Column, double Row, CString szStr)
{
	if(NULL == m_pdispWorksheets[sheet])
		return FALSE;
	if(szStr.IsEmpty())
		return FALSE;
	long nBuffSize = szStr.GetLength();
	

	VARIANTARG vargRng;
	
	ClearAllArgs();
	AddArgumentDouble(NULL, 0, Column);
	AddArgumentDouble(NULL, 0, Row);
	if(!ExlInvoke(m_pdispWorksheets[sheet], L"Cells",&vargRng, DISPATCH_PROPERTYGET, DISP_FREEARGS))
		return FALSE;

    AddArgumentCString(NULL, 0, szStr );
	if (!ExlInvoke(vargRng.pdispVal, L"Value", NULL, DISPATCH_PROPERTYPUT, 0))
		return FALSE;
	ReleaseVariant(&vargRng);
	
	
	return TRUE;
}

BOOL CXLAutomation::SetRangeValueDouble(SheetName sheet, LPOLESTR lpszRef, double d)
{
	if(NULL == m_pdispWorksheets[sheet])
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
	if(NULL != m_pdispExcelApp)
	{
		m_pdispExcelApp->Release();
		m_pdispExcelApp = NULL;
	}

	if(NULL != m_pdispWorkbook)
	{
		m_pdispWorkbook->Release();
		m_pdispWorkbook = NULL;
	}

	// Release all worksheet dispatch pointers
	for (int i = 0; i < WS_NUM_SHEET_COUNT; i++) {
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
	if(NULL == m_pdispWorksheets[sheet])
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

//Get Worksheet.Calls(nColumn, nRow).Value
//This method is not fully tested - see code coments 
CString CXLAutomation::GetCellValueCString(SheetName sheet, int nColumn, int nRow)
{
	CString szValue =_T("");
	if(NULL == m_pdispWorksheets[sheet])
		return szValue;
	
	VARIANTARG vargRng, vargValue;
	
	ClearAllArgs();
	AddArgumentDouble(NULL, 0, nColumn);
	AddArgumentDouble(NULL, 0, nRow);
	if(!ExlInvoke(m_pdispWorksheets[sheet], L"Cells",&vargRng, DISPATCH_PROPERTYGET, DISP_FREEARGS))
		return szValue;
    
	if (!ExlInvoke(vargRng.pdispVal, L"Value", &vargValue, DISPATCH_PROPERTYGET, 0))
		return szValue;

	VARTYPE Type = vargValue.vt;
	switch (Type) 
		{
			case VT_UI1:
				{
					unsigned char nChr = vargValue.bVal;
					//szValue = nChr;
					szValue.Format(_T("%d"),nChr);// = nChr;
				}
				break;
			case VT_I4:
				{
					long nVal = vargValue.lVal;
					szValue.Format(_T("%i"), nVal);
				}
				break;
			case VT_R4:
				{
					float fVal = vargValue.fltVal;
					szValue.Format(_T("%f"), fVal);
				}
				break;
			case VT_R8:
				{
					double dVal = vargValue.dblVal;
					szValue.Format(_T("%f"), dVal);
				}
				break;
			case VT_BSTR:
				{
					BSTR b = vargValue.bstrVal;
					szValue = b;
				}
				break;
			case VT_BYREF|VT_UI1:
				{
					//Not tested
					unsigned char* pChr = vargValue.pbVal;
//					szValue = *pChr;
					szValue.Format(_T("%f"), *pChr);
				}
				break;
			case VT_BYREF|VT_BSTR:
				{
					//Not tested
					BSTR* pb = vargValue.pbstrVal;
					szValue = *pb;
				}
			case 0:
				{
					//Empty
					szValue = _T("");
				}

				break;
		}
	
		
//	ReleaseVariant(&vargRng);
//	ReleaseVariant(&vargValue);
	
	return szValue;

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
	for (int i = 0; i < WS_NUM_SHEET_COUNT; i++) {
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

