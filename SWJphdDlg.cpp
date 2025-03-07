
// SWJphdDlg.cpp: 구현 파일
//

#include "pch.h"
#include "GlobalEnv.h"
#include "framework.h"
#include "SWJphd.h"
#include "SWJphdDlg.h"
#include "DlgProxy.h"
#include "afxdialogex.h"
#include "XLEzAutomation.h"

// Sheet enumeration for easy reference
enum SheetName {
	WS_NUM_GENV = 0,
	WS_NUM_DASHBOARD,
	WS_NUM_PROJECT,
	WS_NUM_ACTIVITY_STRUCT,
	WS_NUM_DEBUG_INFO,
	WS_TOTAL_SHEET_COUNT // Total number of sheets
};

LPOLESTR gSheetsName[WS_TOTAL_SHEET_COUNT] = { L"GlobalEnv", L"dashboard", L"project", L"activity_struct",L"Debug_info" };


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CSWJphdDlg 대화 상자


IMPLEMENT_DYNAMIC(CSWJphdDlg, CDialogEx);

CSWJphdDlg::CSWJphdDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_SWJPHD_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = nullptr;
	memset(&m_gEnv, 0, sizeof(GLOBAL_ENV));
}

CSWJphdDlg::~CSWJphdDlg()
{
	// 이 대화 상자에 대한 자동화 프록시가 있을 경우 이 대화 상자에 대한
	//  후방 포인터를 null로 설정하여
	//  대화 상자가 삭제되었음을 알 수 있게 합니다.
	if (m_pAutoProxy != nullptr)
		m_pAutoProxy->m_pDialog = nullptr;
}

void CSWJphdDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CSWJphdDlg, CDialogEx)
	ON_WM_CLOSE()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BTN_ENV_OPEN, &CSWJphdDlg::OnBnClickedBtnEnvOpen)
	ON_BN_CLICKED(IDC_BTN_SAVE_FILE_NAME, &CSWJphdDlg::OnBnClickedBtnSaveFileName)
	ON_BN_CLICKED(IDC_BTN_ENV_LOAD, &CSWJphdDlg::OnBnClickedBtnEnvLoad)
	ON_BN_CLICKED(IDC_BUTTON1, &CSWJphdDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// CSWJphdDlg 메시지 처리기

BOOL CSWJphdDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 이 대화 상자의 아이콘을 설정합니다.  응용 프로그램의 주 창이 대화 상자가 아닐 경우에는
	//  프레임워크가 이 작업을 자동으로 수행합니다.
	SetIcon(m_hIcon, TRUE);			// 큰 아이콘을 설정합니다.
	SetIcon(m_hIcon, FALSE);		// 작은 아이콘을 설정합니다.

	// TODO: 여기에 추가 초기화 작업을 추가합니다.
		
	srand((unsigned)time(NULL)); // 무작위 시드 초기화
	// 컨트롤 변수와 연결
	m_editEnvFilePath.SubclassDlgItem(IDC_EDIT_ENV_FILE_PATH, this);
	m_editSaveFilePath.SubclassDlgItem(IDC_EDIT_SAVE_FILE_PATH, this);

	// 레지스트리에서 파일 경로 읽기
	LoadFilePathFromRegistry();


	return TRUE;  // 포커스를 컨트롤에 설정하지 않으면 TRUE를 반환합니다.
}

// 대화 상자에 최소화 단추를 추가할 경우 아이콘을 그리려면
//  아래 코드가 필요합니다.  문서/뷰 모델을 사용하는 MFC 애플리케이션의 경우에는
//  프레임워크에서 이 작업을 자동으로 수행합니다.

void CSWJphdDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 그리기를 위한 디바이스 컨텍스트입니다.

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 클라이언트 사각형에서 아이콘을 가운데에 맞춥니다.
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 아이콘을 그립니다.
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// 사용자가 최소화된 창을 끄는 동안에 커서가 표시되도록 시스템에서
//  이 함수를 호출합니다.
HCURSOR CSWJphdDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// 컨트롤러에서 해당 개체 중 하나를 계속 사용하고 있을 경우
//  사용자가 UI를 닫을 때 자동화 서버를 종료하면 안 됩니다.  이들
//  메시지 처리기는 프록시가 아직 사용 중인 경우 UI는 숨기지만,
//  UI가 표시되지 않아도 대화 상자는
//  남겨 둡니다.

void CSWJphdDlg::OnClose()
{
	// 레지스트리에 파일 경로 저장
	SaveFilePathToRegistry();
	if (CanExit())
		CDialogEx::OnClose();
}

void CSWJphdDlg::OnOK()
{
	SaveFilePathToRegistry();
	if (CanExit())
		CDialogEx::OnOK();
}

void CSWJphdDlg::OnCancel()
{
	SaveFilePathToRegistry();
	if (CanExit())
		CDialogEx::OnCancel();
}

BOOL CSWJphdDlg::CanExit()
{
	// 프록시 개체가 계속 남아 있으면 자동화 컨트롤러에서는
	//  이 애플리케이션을 계속 사용합니다.  대화 상자는 남겨 두지만
	//  해당 UI는 숨깁니다.
	if (m_pAutoProxy != nullptr)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}

	return TRUE;
}


void CSWJphdDlg::LoadFilePathFromRegistry()
{
	CRegKey regKey;
	LONG lResult = regKey.Open(HKEY_CURRENT_USER, _T("Software\\SWJphd"), KEY_READ);

	if (lResult == ERROR_SUCCESS) {
		TCHAR szEnvValue[MAX_PATH] = { 0 };
		ULONG nChars = MAX_PATH;
		if (regKey.QueryStringValue(_T("LastEnvFilePath"), szEnvValue, &nChars) == ERROR_SUCCESS) {
			m_strEnvFilePath = szEnvValue;
		}

		TCHAR szSaveValue[MAX_PATH] = { 0 };
		nChars = MAX_PATH;
		if (regKey.QueryStringValue(_T("LastSaveFilePath"), szSaveValue, &nChars) == ERROR_SUCCESS) {
			m_strSaveFilePath = szSaveValue;
		}

		// 명시적으로 레지스트리 닫기 (필수는 아님)
		regKey.Close();
	}

	// 에디트 컨트롤에 값 설정
	if (GetDlgItem(IDC_EDIT_ENV_FILE_PATH)) {
		SetDlgItemText(IDC_EDIT_ENV_FILE_PATH, m_strEnvFilePath);
	}
	if (GetDlgItem(IDC_EDIT_SAVE_FILE_PATH)) {
		SetDlgItemText(IDC_EDIT_SAVE_FILE_PATH, m_strSaveFilePath);
	}
}

void CSWJphdDlg::SaveFilePathToRegistry()
{
	CRegKey regKey;

	// 레지스트리 키 생성 (또는 열기)
	if (regKey.Create(HKEY_CURRENT_USER, _T("Software\\SWJphd")) == ERROR_SUCCESS) {
		// 에디트 컨트롤에서 값을 가져와 저장
		GetDlgItemText(IDC_EDIT_ENV_FILE_PATH, m_strEnvFilePath);
		regKey.SetStringValue(_T("LastEnvFilePath"), m_strEnvFilePath);

		GetDlgItemText(IDC_EDIT_SAVE_FILE_PATH, m_strSaveFilePath);
		regKey.SetStringValue(_T("LastSaveFilePath"), m_strSaveFilePath);
	}
}

void CSWJphdDlg::OnBnClickedBtnEnvOpen()
{
	// 파일 선택 다이얼로그 필터 설정 (xlsx 파일만)
	CFileDialog fileDlg(TRUE, _T("xlsx"), nullptr, OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST,
		_T("Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*||"), this);

	// 다이얼로그 표시
	if (fileDlg.DoModal() == IDOK)
	{
		// 선택한 파일 경로 가져오기
		m_strEnvFilePath = fileDlg.GetPathName();

		// 에디트 박스에 경로 설정
		SetDlgItemText(IDC_EDIT_ENV_FILE_PATH, m_strEnvFilePath);
	}
}

void CSWJphdDlg::OnBnClickedBtnSaveFileName()
{
	// 파일 선택 다이얼로그 필터 설정 (xlsx 파일만)
	CFileDialog fileDlg(TRUE, _T("xlsx"), nullptr, OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST,
		_T("Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*||"), this);

	// 다이얼로그 표시
	if (fileDlg.DoModal() == IDOK)
	{
		// 선택한 파일 경로 가져오기
		m_strSaveFilePath = fileDlg.GetPathName();

		// 에디트 박스에 경로 설정
		SetDlgItemText(IDC_EDIT_SAVE_FILE_PATH, m_strSaveFilePath);
	}
}

// 환경변수들을 엑셀파일에서 가져온다.
void CSWJphdDlg::OnBnClickedBtnEnvLoad()
{
	if (m_strEnvFilePath.IsEmpty()) {
		AfxMessageBox(_T("파일 경로가 설정되지 않았습니다."));
		return;
	}

	// 엑셀 파일 열기
	CXLEzAutomation xlAuto;
	if (!xlAuto.OpenExcelFile(m_strEnvFilePath, gSheetsName, WS_TOTAL_SHEET_COUNT)) {
		AfxMessageBox(_T("엑셀 파일을 열 수 없습니다."));
		return ;
	}

	int i = 7;	// Global 환경변수 가져오기
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.simulationPeriod);	i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.maxPeriod);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.technicalFee);	i++;

	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.higHrCount);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.midHrCount);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.lowHrCount);		i++;

	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.higHrCost);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.midHrCost);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.lowHrCost);		i++;

	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.initialFunds);	i++;

	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.extPrjInTime);	i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.intPrjInTime);	i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.minDuration);	i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.maxDuration);	i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.durRule);		i++;

	// mode의 환경변수 가져오기
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.minMode);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.maxMode);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.lifeCycle);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.profitRate);	    i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.mu0Rate);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.sigma0Rate);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.mu1Rate);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.mu2Rate);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.sigma1Rate);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.sigma2Rate);		i++;

	// 엑셀 파일 닫기
	xlAuto.ReleaseExcel();

	// Global 환경 변수 업데이트
	SetDlgItemInt(IDC_EDIT_SIM_PERIOD,	m_gEnv.simulationPeriod);
	SetDlgItemInt(IDC_EDIT_HIG_HR,		m_gEnv.higHrCount);
	SetDlgItemInt(IDC_EDIT_MID_HR,		m_gEnv.midHrCount);
	SetDlgItemInt(IDC_EDIT_LOW_HR,		m_gEnv.lowHrCount);
	SetDlgItemInt(IDC_EDIT_INI_FUNDS,	m_gEnv.initialFunds);

	CString strTemp;
	strTemp.Format(_T("%.2f"), m_gEnv.extPrjInTime);  //
	SetDlgItemText(IDC_EDIT_AVG_PROJECT, strTemp);

	//SetDlgItemInt(IDC_EDIT_AVG_PROJECT,	m_gEnv.avgWeeklyProjects,FALSE);
	//SetDlgItemInt(IDC_EDIT_INI_FUNDS,	m_gEnv.higHrCostRate);
	//SetDlgItemInt(IDC_EDIT_INI_FUNDS,	m_gEnv.midHrCostRate);
	//SetDlgItemInt(IDC_EDIT_INI_FUNDS,	m_gEnv.lowHrCostRate);

	m_pCreator.Init(m_strEnvFilePath, &m_gEnv);
	PrintAllProject(&m_pCreator);

}

void CSWJphdDlg::OnBnClickedButton1()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.

}
void CSWJphdDlg::PrintAllProject(CCreator* pCreator)
{
	// 엑셀 파일 열기
	CXLEzAutomation xlAuto;
    LPOLESTR sheetsName[1] = { L"project"};
	if (!xlAuto.OpenExcelFile(m_strSaveFilePath, sheetsName,1) )
	{
		AfxMessageBox(_T("엑셀 파일을 열 수 없습니다."));
		return;
	}

#define _PRINT_WIDTH 35
	int nPrjCnt = 0;
	nPrjCnt = pCreator->m_totalProjectNum;
    CString strTitle[1][_PRINT_WIDTH] = {
            _T("category"), _T("ID"), _T("발주일"), _T("시작가능"),_T("기간"), _T("시작"), _T("끝"), 
			_T("수익금"),_T("선금"),_T("중도금"),_T("잔금"),_T("선금일"), _T("중도금일"),_T("잔금일"),
            _T("고급"),_T("중급"),_T("초급"),_T("lifeCycle"),_T("MU"),_T("SIGMA"),_T("고정수익"),
			_T("고급"),_T("중급"),_T("초급"),_T("lifeCycle"),_T("MU"),_T("SIGMA"),_T("고정수익"),
			_T("고급"),_T("중급"),_T("초급"),_T("lifeCycle"),_T("MU"),_T("SIGMA"),_T("고정수익")
    };
    xlAuto.WriteArrayToRange(0, 1, 1, (CString*)strTitle, 1, _PRINT_WIDTH);
    
    for (int i =0; i <nPrjCnt; i++)
    {
        PROJECT* pProject = &(pCreator->m_pProjects[0][i]);
        int j = 1 ;
        xlAuto.SetCellValue(0, i + 2, j, pProject->category);			j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->ID);					j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->createTime);			j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->startAbleTime);		j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->duration);			j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->startTime);			j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->endTime);			j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->revenue);			j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->firstPay);			j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->secondPay);			j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->finalPay);			j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->firstPayTime);		j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->secondPayTime);		j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->finalPayTime);		j++;

        xlAuto.SetCellValue(0, i + 2, j, pProject->mode0.higHrCount);	j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->mode0.midHrCount);	j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->mode0.lowHrCount);	j++;        
        xlAuto.SetCellValue(0, i + 2, j, pProject->mode0.lifeCycle);	j++;        
        xlAuto.SetCellValue(0, i + 2, j, pProject->mode0.mu);			j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->mode0.sigma);		j++;
		xlAuto.SetCellValue(0, i + 2, j, pProject->mode0.fixedIncome);	j++;

        xlAuto.SetCellValue(0, i + 2, j, pProject->mode1.higHrCount);	j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->mode1.midHrCount);	j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->mode1.lowHrCount);	j++;        
        xlAuto.SetCellValue(0, i + 2, j, pProject->mode1.lifeCycle);	j++;        
        xlAuto.SetCellValue(0, i + 2, j, pProject->mode1.mu);			j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->mode1.sigma);		j++;
		xlAuto.SetCellValue(0, i + 2, j, pProject->mode1.fixedIncome);	j++;

        xlAuto.SetCellValue(0, i + 2, j, pProject->mode2.higHrCount);	j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->mode2.midHrCount);	j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->mode2.lowHrCount);	j++;        
        xlAuto.SetCellValue(0, i + 2, j, pProject->mode2.lifeCycle);	j++;        
        xlAuto.SetCellValue(0, i + 2, j, pProject->mode2.mu);			j++;
        xlAuto.SetCellValue(0, i + 2, j, pProject->mode2.sigma);		j++;
		xlAuto.SetCellValue(0, i + 2, j, pProject->mode2.fixedIncome);	j++;
    }

    // 엑셀 파일 닫기
    xlAuto.ReleaseExcel();


}