
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


LPOLESTR gSaveSheets[WS_TOTAL_SHEET_COUNT] = { L"GlobalEnv", L"dashboard", L"project", L"result" };



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

	if (NULL != m_pCreator)	
		delete m_pCreator;

	if (NULL != m_pSaveXl)
		delete m_pSaveXl;

	if (NULL != m_pCompany)
		delete m_pCompany;
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
	ON_BN_CLICKED(IDC_BTN_STEP_BY_STEP, &CSWJphdDlg::OnBnClickedBtnStepByStep)
	ON_BN_CLICKED(IDC_BTN_ENV_SAVE, &CSWJphdDlg::OnBnClickedBtnEnvSave)
	ON_BN_CLICKED(IDC_BTN_GO, &CSWJphdDlg::OnBnClickedBtnGo)
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
	m_pCreator		= NULL;
	m_pSaveXl		= NULL;
	m_pCompany		= NULL;
	m_stepByStepCnt	= -1;
	
	// 버튼 비활성화
	GetDlgItem(IDC_BTN_STEP_BY_STEP)->EnableWindow(FALSE);
	GetDlgItem(IDC_BTN_ENV_SAVE)->EnableWindow(FALSE);

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
	// 파일 선택 다이얼로그 필터 설정 (xlsx, xlsm 파일만)
	CFileDialog fileDlg(TRUE, _T("xlsx"), nullptr, OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST,
		_T("Excel Files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All Files (*.*)|*.*||"), this);

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
	// 파일 선택 다이얼로그 필터 설정 (xlsx, xlsm 파일만)
	CFileDialog fileDlg(TRUE, _T("xlsx"), nullptr, OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST,
		_T("Excel Files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All Files (*.*)|*.*||"), this);

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
	// 파일경로 업데이트
	GetDlgItemText(IDC_EDIT_ENV_FILE_PATH, m_strEnvFilePath);
	GetDlgItemText(IDC_EDIT_SAVE_FILE_PATH, m_strSaveFilePath);

	if (m_strEnvFilePath.IsEmpty()) {
		AfxMessageBox(_T("환경변수 파일 경로가 설정되지 않았습니다."));
		return;
	}

	// 엑셀 파일 열기
	CXLEzAutomation xlAuto;
	//if (!xlAuto.OpenExcelFile(m_strEnvFilePath, gSheetsName, WS_TOTAL_SHEET_COUNT)) {
	CString strSheet = _T("GlobalEnv");
	if (!xlAuto.OpenExcelFile(m_strEnvFilePath, strSheet, FALSE)) {
		AfxMessageBox(_T("환경변수 파일을 열수 없습니다."));
		return ;
	}

	int i = 7;	// Global 환경변수 가져오기
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.problemCount);	i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.selectedMode);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.simulationPeriod);	i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.maxPeriod);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.technicalFee);	i++;

	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.higHrCount);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.midHrCount);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.lowHrCount);		i++;

	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.higHrCost);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.midHrCost);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.lowHrCost);		i++;
	
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.fundsHoldTerm);	i++;
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
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.sigma1Rate);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.mu2Rate);		i++;
	xlAuto.GetCellValue(WS_NUM_GENV, i, 3, &m_gEnv.sigma2Rate);		i++;

	// 엑셀 파일 닫기
	xlAuto.ReleaseExcel();

	// Global 환경 변수 값들을 엑셀파일에 있는 값으로 업데이트 한다.
	CString strTemp;

	SetDlgItemInt(IDC_EDIT_PROBLEM_COUNT, m_gEnv.problemCount);
	SetDlgItemInt(IDC_EDIT_SELECT_MODE, m_gEnv.selectedMode);

	SetDlgItemInt(IDC_EDIT_SIM_PERIOD,	m_gEnv.simulationPeriod);
	strTemp.Format(_T("%.2f"), m_gEnv.technicalFee);  //
	SetDlgItemText(IDC_EDIT_TECH_FEE, strTemp);

	SetDlgItemInt(IDC_EDIT_HIG_HR,		m_gEnv.higHrCount);
	SetDlgItemInt(IDC_EDIT_MID_HR,		m_gEnv.midHrCount);
	SetDlgItemInt(IDC_EDIT_LOW_HR,		m_gEnv.lowHrCount);

	SetDlgItemInt(IDC_EDIT_HIG_COST, m_gEnv.higHrCost);
	SetDlgItemInt(IDC_EDIT_MID_COST, m_gEnv.midHrCost);
	SetDlgItemInt(IDC_EDIT_LOW_COST, m_gEnv.lowHrCost);

	SetDlgItemInt(IDC_EDIT_FUND_RATE, m_gEnv.fundsHoldTerm);
	SetDlgItemInt(IDC_EDIT_INI_FUNDS,	m_gEnv.initialFunds);

	strTemp.Format(_T("%.2f"), m_gEnv.extPrjInTime);
	SetDlgItemText(IDC_EDIT_EX_PROBABILITY, strTemp);
	strTemp.Format(_T("%.2f"), m_gEnv.intPrjInTime);
	SetDlgItemText(IDC_EDIT_IN_PROBABILITY, strTemp);

	SetDlgItemInt(IDC_EDIT_MAX_DURATION, m_gEnv.maxDuration);
	SetDlgItemInt(IDC_EDIT_MIN_DURATION, m_gEnv.minDuration);

	SetDlgItemInt(IDC_EDIT_MAX_MODE, m_gEnv.maxMode);
	SetDlgItemInt(IDC_EDIT_MIN_MODE, m_gEnv.minMode);
	SetDlgItemInt(IDC_EDIT_LIFE_CYCLE, m_gEnv.lifeCycle);
	strTemp.Format(_T("%.2f"), m_gEnv.profitRate);
	SetDlgItemText(IDC_EDIT_PROFIT_RATE, strTemp);

	SetDlgItemInt(IDC_EDIT_MU0, m_gEnv.mu0Rate);
	SetDlgItemInt(IDC_EDIT_SIGMA0, m_gEnv.sigma0Rate);

	strTemp.Format(_T("%.2f"), m_gEnv.mu1Rate);
	SetDlgItemText(IDC_EDIT_MU1, strTemp);

	strTemp.Format(_T("%.2f"), m_gEnv.sigma1Rate);
	SetDlgItemText(IDC_EDIT_SIGMA1, strTemp);

	strTemp.Format(_T("%.2f"), m_gEnv.mu2Rate);
	SetDlgItemText(IDC_EDIT_MU2, strTemp);

	strTemp.Format(_T("%.2f"), m_gEnv.sigma2Rate);
	SetDlgItemText(IDC_EDIT_SIGMA2, strTemp);

	// 버튼 활성화
	GetDlgItem(IDC_BTN_STEP_BY_STEP)->EnableWindow(TRUE);
	GetDlgItem(IDC_BTN_ENV_SAVE)->EnableWindow(TRUE);
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
        //xlAuto.SetCellValue(0, i + 2, j, pProject->revenue);			j++;
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

void CSWJphdDlg::OnBnClickedBtnStepByStep()
{
	// 초기화에 추가??	
	// step 카운터가 0이면 creator과 companny 생성
	if (-1 == m_stepByStepCnt)
	{
		// 다른 버튼들은 disable 하고 step by step 모드로 진입함.
		m_stepByStepCnt = 0;
		ScreenValueToLocalValue();

		if(NULL != m_pSaveXl)
		{
			m_pSaveXl->ReleaseExcel();
			delete m_pSaveXl;
			m_pSaveXl = NULL;
		}
		m_pSaveXl = new CXLEzAutomation;
		if (!m_pSaveXl->CreateExcelFile(m_strSaveFilePath, gSaveSheets, WS_TOTAL_SHEET_COUNT))
		{
			return;
		}
				
		if (NULL != m_pCreator)
		{
			delete m_pCreator;			
			m_pCreator = NULL;
		}
		m_pCreator = new CCreator;
		m_pCreator->Init(&m_gEnv);
		m_pCreator->SetActMode(m_gEnv.selectedMode);

		if (NULL != m_pCompany)
		{
			delete m_pCompany;
			m_pCompany = NULL;
		}
		m_pCompany = new CCompany(&m_gEnv, m_pCreator);
		m_pCompany->Init();		

		m_pSaveXl->ClearSheetContents(WS_NUM_RESULT);

		//m_pSaveXl->OpenExcelFile(m_strSaveFilePath,_T("project"),TRUE);
		m_pSaveXl->ClearSheetContents(WS_NUM_PROJECT);
		PrintProjectSheetHeader(m_pSaveXl, WS_NUM_PROJECT);

		m_pSaveXl->ClearSheetContents(WS_NUM_DASHBOARD);
		PrintDBoardHeader(m_pSaveXl, WS_NUM_DASHBOARD);
		LoacalValueToExcel(m_pSaveXl, WS_NUM_GENV, m_pCreator->m_pEnv,FALSE);// 환경 변수 저장
	}

	// 기간이 다 될때 까지
	if(m_stepByStepCnt < m_gEnv.maxPeriod)
	{
		if (FALSE == m_pCompany->Decision(m_stepByStepCnt))  // 이번 스탭의 기간에 결정해야 할 일들			
		{
			GetDlgItem(IDC_BTN_ENV_SAVE)->EnableWindow(TRUE);
			GetDlgItem(IDC_BTN_STEP_BY_STEP)->EnableWindow(FALSE);

			PrintOneTime(m_pSaveXl, m_pCreator, m_pCompany, m_stepByStepCnt);

			AfxMessageBox(_T("파산"));

			m_stepByStepCnt = -1;

			return;
		}
	}

	// 진행한 한주간의 정보 기록
	PrintOneTime(m_pSaveXl, m_pCreator, m_pCompany, m_stepByStepCnt);

	// 기간이 다 되었으면 
	if (m_pCompany->m_pEnv->maxPeriod <= m_stepByStepCnt)
	{	
		m_stepByStepCnt = -1;
		GetDlgItem(IDC_BTN_ENV_SAVE)->EnableWindow(TRUE);
		GetDlgItem(IDC_BTN_STEP_BY_STEP)->EnableWindow(FALSE);
		return;
	}
	m_stepByStepCnt ++;
}

void CSWJphdDlg::ScreenValueToLocalValue()
{
	CString strTemp;

	m_gEnv.problemCount = GetDlgItemInt(IDC_EDIT_PROBLEM_COUNT);
	m_gEnv.selectedMode = GetDlgItemInt(IDC_EDIT_SELECT_MODE);
	m_gEnv.simulationPeriod = GetDlgItemInt(IDC_EDIT_SIM_PERIOD);

	GetDlgItemText(IDC_EDIT_TECH_FEE, strTemp);
	m_gEnv.technicalFee = _tstof(strTemp);

	m_gEnv.higHrCount = GetDlgItemInt(IDC_EDIT_HIG_HR);
	m_gEnv.midHrCount = GetDlgItemInt(IDC_EDIT_MID_HR);
	m_gEnv.lowHrCount = GetDlgItemInt(IDC_EDIT_LOW_HR);

	m_gEnv.higHrCost = GetDlgItemInt(IDC_EDIT_HIG_COST);
	m_gEnv.midHrCost = GetDlgItemInt(IDC_EDIT_MID_COST);
	m_gEnv.lowHrCost = GetDlgItemInt(IDC_EDIT_LOW_COST);

	m_gEnv.fundsHoldTerm = GetDlgItemInt(IDC_EDIT_FUND_RATE);
	m_gEnv.initialFunds = GetDlgItemInt(IDC_EDIT_INI_FUNDS);

	GetDlgItemText(IDC_EDIT_EX_PROBABILITY, strTemp);
	m_gEnv.extPrjInTime = _tstof(strTemp);

	GetDlgItemText(IDC_EDIT_IN_PROBABILITY, strTemp);
	m_gEnv.intPrjInTime = _tstof(strTemp);

	m_gEnv.maxDuration = GetDlgItemInt(IDC_EDIT_MAX_DURATION);
	m_gEnv.minDuration = GetDlgItemInt(IDC_EDIT_MIN_DURATION);

	m_gEnv.maxMode = GetDlgItemInt(IDC_EDIT_MAX_MODE);
	m_gEnv.minMode = GetDlgItemInt(IDC_EDIT_MIN_MODE);
	m_gEnv.lifeCycle = GetDlgItemInt(IDC_EDIT_LIFE_CYCLE);

	GetDlgItemText(IDC_EDIT_PROFIT_RATE, strTemp);
	m_gEnv.profitRate = _tstof(strTemp);

	m_gEnv.mu0Rate = GetDlgItemInt(IDC_EDIT_MU0);
	m_gEnv.sigma0Rate = GetDlgItemInt(IDC_EDIT_SIGMA0);

	GetDlgItemText(IDC_EDIT_MU1, strTemp);
	m_gEnv.mu1Rate = _tstof(strTemp);

	GetDlgItemText(IDC_EDIT_SIGMA1, strTemp);
	m_gEnv.sigma1Rate = _tstof(strTemp);

	GetDlgItemText(IDC_EDIT_MU2, strTemp);
	m_gEnv.mu2Rate = _tstof(strTemp);

	GetDlgItemText(IDC_EDIT_SIGMA2, strTemp);
	m_gEnv.sigma2Rate = _tstof(strTemp);


	m_gEnv.maxPeriod = m_gEnv.simulationPeriod+24 ;
	SetDlgItemInt(IDC_EDIT_INI_FUNDS, m_gEnv.maxPeriod);

	m_gEnv.initialFunds = (m_gEnv.higHrCount * m_gEnv.higHrCost + m_gEnv.midHrCount * m_gEnv.midHrCost + m_gEnv.lowHrCount * m_gEnv.lowHrCost) * m_gEnv.fundsHoldTerm;
	SetDlgItemInt(IDC_EDIT_INI_FUNDS, m_gEnv.initialFunds);

	return;
}

void CSWJphdDlg::OnBnClickedBtnEnvSave()
{
	// Global 환경 변수 값들을 엑셀파일로 저장한다.
	ScreenValueToLocalValue();

	// 엑셀 파일 열기
	CXLEzAutomation xlGEnv;	
	CString strSheet = _T("GlobalEnv");
	if (!xlGEnv.OpenExcelFile(m_strEnvFilePath, strSheet, FALSE)) {
		AfxMessageBox(_T("환경변수 파일을 열수 없습니다."));
		return;
	}

	LoacalValueToExcel(&xlGEnv, 0, &m_gEnv, TRUE);

	xlGEnv.GetCellValue(WS_NUM_GENV, 8, 3, &m_gEnv.maxPeriod);
	SetDlgItemInt(IDC_EDIT_INI_FUNDS, m_gEnv.maxPeriod);

	xlGEnv.GetCellValue(WS_NUM_GENV, 17, 3, &m_gEnv.initialFunds);
	SetDlgItemInt(IDC_EDIT_INI_FUNDS, m_gEnv.initialFunds);
	xlGEnv.SaveAndCloseExcelFile(m_strEnvFilePath);
	xlGEnv.ReleaseExcel();


	if (NULL != m_pCreator ){
		delete m_pCreator;
		m_pCreator = NULL;
	}

	if (NULL != m_pSaveXl) {
		delete m_pSaveXl;
		m_pSaveXl = NULL;
	}

	if (NULL != m_pCompany) {
		delete m_pCompany;
		m_pCompany = NULL;
	}

	m_stepByStepCnt = -1;
	//GetDlgItem(IDC_BTN_STEP_BY_STEP)->EnableWindow(FALSE);
}


void CSWJphdDlg::OnBnClickedBtnGo()
{	
	// 결과를 담을 배열
	ScreenValueToLocalValue();

	Dynamic2DArray resultArr;
	int rows, cols;
	rows = m_gEnv.problemCount;
	// 순번,  수익0,수익1,수익2,  완료일0,완료일1,완료일2,  계급구간,계급0,계급1,계급2, 누적0,누적1,누적2
	cols = 14;
	resultArr.Resize(rows, cols);
	for (int i = 0; i < rows; ++i) {
		resultArr[i][0] = i + 1;
	}



	if (NULL != m_pSaveXl)	{
		m_pSaveXl->ReleaseExcel();
		delete m_pSaveXl;
	}

	m_pSaveXl = new CXLEzAutomation;

	if (!m_pSaveXl->CreateExcelFile(m_strSaveFilePath, gSaveSheets, WS_TOTAL_SHEET_COUNT))
	{
		AfxMessageBox(_T("error"));
		return ;
	}
	LoacalValueToExcel(m_pSaveXl, WS_NUM_GENV, &m_gEnv, FALSE);// 환경 변수 저장

	for(int i = 0; i < rows ;i++) // 실행 횟수
	{
		// 동일한 조건에서 모드만 다르게 해서 row 만큼 실행하고 결과를 얻는다.
		if (FALSE == InitCompanyAndCreator())
		{
			AfxMessageBox(_T("InitCompanyAndCreator 에러"));
			return;
		}

		for (int j=0 ;j <3; j++) // 모드별로 진행
		{	
			m_pCompany->m_pCreator->SetActMode(j);// 변경된 모드로 Project 변경
			m_pCompany->Init(); // 변경된 모드로 Company 초기화

			_RESULT_ simResult;
			simResult.time		= 0;
			simResult.balance	= 0;
			RunSimulator(&simResult);// 시뮬레이션 실행하고 결과 리턴

			// 결과 기록
			resultArr[i][j+1] = simResult.balance;
			resultArr[i][j+4] = simResult.time;
		}
	}

	// 0:순번,  수익0,수익1,수익2,  완료일0,완료일1,완료일2,  계급구간,계급0,계급1,계급2, 누적구간,누적0,누적1,누적2
	MakeFreq(&resultArr,0, 1,rows, 3, 0, 7, 50);

	int totalSize = rows * cols;  // Total number of elements	
	int* tempBuf = new int[totalSize];  // Allocate memory for the total number of elements
	resultArr.copyToContinuousMemory(tempBuf, totalSize);

	PrintResultHeader(m_pSaveXl, WS_NUM_RESULT);
	m_pSaveXl->WriteArrayToRange(WS_NUM_RESULT, 3,1,tempBuf, rows, cols);

	delete[] tempBuf;
}

void CSWJphdDlg::RunSimulator(_RESULT_* simResult)
{
	// 마지막 기간의 결과 확인은 기간이 다 되고난 다음 시점까지 진행해야 나옴
	int thisTime;
	for (thisTime = 0; thisTime <= m_gEnv.simulationPeriod; thisTime++)
	{
		if (FALSE == m_pCompany->Decision(thisTime))  // 이번 스탭의 기간에 결정해야 할 일들			
		{
			simResult->time		= thisTime;
			simResult->balance	= m_pCompany->m_balanceTable[0][thisTime-1];
			return;
		}
	}
	
	simResult->time = thisTime-1;
	simResult->balance = m_pCompany->m_balanceTable[0][thisTime - 2];
	return;	
}

BOOL CSWJphdDlg::InitCompanyAndCreator()
{
    if (NULL != m_pCreator)
    {
        delete m_pCreator;        
        m_pCreator = NULL;
    }
    m_pCreator = new CCreator;
    m_pCreator->Init(&m_gEnv);
    m_pCreator->SetActMode(m_gEnv.selectedMode);

    if (NULL != m_pCompany)
    {
        delete m_pCompany;
        m_pCompany = NULL;
    }

    m_pCompany = new CCompany(&m_gEnv, m_pCreator);
    m_pCompany->Init();

	return TRUE;
}