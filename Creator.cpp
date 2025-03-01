#include "pch.h"
#include "GlobalEnv.h"
#include "Creator.h"
#include <cctype>   // toupper 함수를 사용하기 위해 필요

CCreator::CCreator()
{	
}
CCreator::CCreator(int count)
{	
}

CCreator::~CCreator()
{	
}

// 문제를 생성한다. 글로벌 환경에서 발생 할 수 있는 프로젝트들을 발생 시키고 파일로 저장한다.
BOOL CCreator::Init(GLOBAL_ENV* pGlobalEnv)
{	
	*(&m_GlobalEnv) = *pGlobalEnv;
	m_pProjects.Resize(0,10);
	m_totalProjectNum = CreateAllProjects();

	return TRUE;
}


/*
1) 월간 내부 프로젝트 발생 확율과 외부 프로젝트 발생 확율을 다르게 한다.
2) 
*/
int CCreator::CreateAllProjects()
{
	int prjectId = 0;
	for (int week = 0; week < m_GlobalEnv.maxWeek; week++)
	{
		if (0 == week % (52)) // 1년에 하나씩 내부 프로젝트 발생
		{
			CreateInternalProject(prjectId++,week);
		}

		int newCnt = PoissonRandom(m_GlobalEnv.avgWeeklyProjects);	// 이번주 발생하는 프로젝트 갯수

		for (int i = 0; i < newCnt; i++) // 외부 프로젝트 발생
		{
			CraterExternalProject(prjectId++,week);
		}
	}

	return prjectId;
}

int CCreator::CreateInternalProject(int Id,int time)
{
	PROJECT Project;
	
	memset(&Project, 0, sizeof(struct PROJECT));

	Project.category		= 1;	// 프로젝트 분류 (0: 외부 / 1: 내부)
	Project.ID				= Id;	// 프로젝트의 번호	
	Project.createTime		= time;	// 발주일
	Project.startAbleTime	= time;	// 시작 가능일. 내부는 바로 진행가능	

	//Project.profit = CalculateHRAndProfit(&Project); // 총 수익을 계산한다.
	//Project.profit = Project.profit * m_GlobalEnv.multiples/(52*3);
	m_pProjects[0][Id] = Project;

	return 0;
}

int CCreator::CraterExternalProject(int Id, int time)
{
	PROJECT Project;

	memset(&Project, 0, sizeof(struct PROJECT));

	Project.category		= 0;	// 프로젝트 분류 (0: 외부 / 1: 내부)
	Project.ID				= Id;	// 프로젝트의 번호	
	Project.createTime		= time;	// 발주일
	Project.startAbleTime	= time + (rand() % 4);  // // 시작 가능일 ( 0에서 3 사이의 정수 난수 생성)
	
	m_pProjects[0][Id] = Project;
		
	return 0;
}

// 활동별 투입 인력 생성 및 프로젝트 전체 기대 수익 계산 함수
double CCreator::CalculateHRAndProfit(PROJECT* pProject) {
	int high = 0, mid = 0, low = 0;

	return CalculateTotalLaborCost(high, mid, low);
}

// 등급별 투입 인력 계산 및 프로젝트의 수익 생성 함수
double CCreator::CalculateTotalLaborCost(int highCount, int midCount, int lowCount) {
	double highLaborCost	= CalculateLaborCost("H") * highCount;
	double midLaborCost	= CalculateLaborCost("M") * midCount;
	double lowLaborCost	= CalculateLaborCost("L") * lowCount;
	
	double totalLaborCost = highLaborCost + midLaborCost + lowLaborCost;
	return totalLaborCost;
}

// 등급별 투입 인력에 따른 수익 계산 함수
// 수정시 다음도 수정 필요 BOOL CCompany::Init(PGLOBAL_ENV pGlobalEnv, int Id, BOOL shouldLoad)
double CCreator::CalculateLaborCost(const std::string& grade) {
	double directLaborCost	= 0;
	double overheadCost	= 0;
	double technicalFee	= 0;
	double totalLaborCost	= 0;

	// 입력된 grade를 대문자로 변환
	char upperGrade = std::toupper(static_cast<unsigned char>(grade[0]));

	switch (upperGrade) {
	case 'H':
		directLaborCost = HI_HR_COST;
		break;
	case 'M':
		directLaborCost = MI_HR_COST;
		break;
	case 'L':
		directLaborCost = LO_HR_COST;
		break;
	default:
		AfxMessageBox(_T("잘못된 등급입니다. 'H', 'M', 'L' 중 하나를 입력하세요."), MB_OK | MB_ICONERROR);
		return 0; // 잘못된 입력 시 함수 종료
	}

	overheadCost = directLaborCost * 1.4;// 0.6; // 간접 비용 계산
	technicalFee = (directLaborCost + overheadCost) * 0.4;// 0.2; // 기술 비용 계산
	totalLaborCost = directLaborCost + overheadCost + technicalFee; // 총 인건비 계산

	return totalLaborCost;
}


// 대금 지급 조건 생성 함수
void CCreator::CalculatePaymentSchedule(PROJECT* pProject) {

	int paymentType;
	int paymentRatio;
	int totalPayments;

	pProject->firstPayTime = 1;
	
	// 6주 이하의 짧은 프로젝트는 선금, 잔금만 있다.
	if (pProject->duration <= 6) {
		paymentType = rand() % 3 + 1;  // 1에서 3 사이의 난수 생성

		switch (paymentType) {
		case 1:
			pProject->firstPay = (int)ceil((double)pProject->profit * 0.3);			
			break;
		case 2:
			pProject->firstPay = (int)ceil((double)pProject->profit * 0.4);
			break;
		case 3:
			pProject->firstPay = (int)ceil((double)pProject->profit * 0.5);
			break;
		}

		pProject->secondPay = pProject->profit - pProject->firstPay;
		totalPayments = 2;
		pProject->secondPayTime = pProject->duration;
	}

	// 7~12주 사이의 프로젝트는 3회에 걸처셔 받는다.
	else if (pProject->duration <= 12) {
		paymentType = rand() % 10 + 1;  // 1에서 10 사이의 난수 생성

		if (paymentType <= 3) {
			paymentRatio = rand() % 3 + 1;  // 1에서 3 사이의 난수 생성

			switch (paymentRatio) {
			case 1:
				pProject->firstPay = (int)ceil((double)pProject->profit * 0.3);
				break;
			case 2:
				pProject->firstPay = (int)ceil((double)pProject->profit * 0.4);
				break;
			case 3:
				pProject->firstPay = (int)ceil((double)pProject->profit * 0.5);
				break;
			}

			pProject->secondPay = pProject->profit - pProject->firstPay;
			totalPayments = 2;
			pProject->secondPayTime = pProject->duration;
		}
		else {
			paymentRatio = rand() % 10 + 1;  // 1에서 10 사이의 난수 생성

			if (paymentRatio <= 6) {
				pProject->firstPay = (int)ceil((double)pProject->profit * 0.3);
				pProject->secondPay = (int)ceil((double)pProject->profit * 0.3);
			}
			else {
				pProject->firstPay = (int)ceil((double)pProject->profit * 0.3);
				pProject->secondPay = (int)ceil((double)pProject->profit * 0.4);
			}

			pProject->finalPay = pProject->profit - pProject->firstPay - pProject->secondPay;
			totalPayments = 3;
			pProject->secondPayTime = (int)ceil((double)pProject->duration / 2);
			pProject->finalPayTime = pProject->duration;
		}
	}

	// 1년 이상의 프로젝트는 3회에 걸처서 받는다.
	else {
		pProject->firstPay = (int)ceil((double)pProject->profit * 0.3);
		pProject->secondPay = (int)ceil((double)pProject->profit * 0.4);
		pProject->finalPay = pProject->profit - pProject->firstPay - pProject->secondPay;

		totalPayments = 3;
		pProject->secondPayTime = (int)ceil((double)pProject->duration / 2);
		pProject->finalPayTime = pProject->duration;
	}
	pProject->nCashFlows = totalPayments;
}

/*
void CCreator::Save(CString filename,CString strInSheetName)
{
	Book* book = xlCreateXMLBook();  // Use xlCreateBook() for xls format	

	// 정품 키 값이 들어 있다. 공개하는 프로젝트에는 포함되어 있지 않다. 
	// 정품 키가 없으면 읽기가 300 컬럼으로 제한된다.
#ifdef INCLUDE_LIBXL_KET
	book->setKey(_LIBXL_NAME, _LIBXL_KEY);
#endif

	Sheet* projectSheet = nullptr;
	Sheet* dashboardSheet = nullptr;

	if (book->load(filename)) {
		// File exists, check for specific sheets

		for (int i = 0; i < book->sheetCount(); ++i) {
			Sheet* sheet = book->getSheet(i);
			if (std::wcscmp(sheet->name(), strInSheetName) == 0) {
				projectSheet = sheet;
				clearSheet(projectSheet);  // Assuming you have a clearSheet function defined
			}
		}

		if (!projectSheet) {
			projectSheet = book->addSheet(strInSheetName);
		}
	}
	else {
		// File does not exist, create new file with sheets
		projectSheet = book->addSheet(strInSheetName);  // Add and assign the 'project' sheet
		//dashboardSheet = book->addSheet(L"dashboard");  // Add and assign the 'dashboard' sheet
	}

	write_global_env(book, projectSheet,&m_GlobalEnv);
	write_project_header(book, projectSheet);

	for (int i = 0; i < m_totalProjectNum; i++) {
		write_project_body(book, projectSheet, &(m_pProjects[0][i]));  // Assuming write_project_body is defined
	}

	projectSheet->writeStr(1, 2, L"TotalPrjNum");
	projectSheet->writeNum(1, 3, m_totalProjectNum);

	// Save and release
	book->save(filename);
	book->release();
}
*/
/*
void CCreator::PrintProjectInfo() {

	Book* book = xlCreateXMLBook();  // Use xlCreateBook() for xls format	

	// 정품 키 값이 들어 있다. 공개하는 프로젝트 소스코드에는 포함되어 있지 않다. 
	// 정품 키가 없으면 읽기가 300 컬럼으로 제한된다. 필요시 구매해서 사용 바람.
	#ifdef INCLUDE_LIBXL_KET
		book->setKey(_LIBXL_NAME, _LIBXL_KEY);
	#endif

	Sheet* projectSheet = nullptr;
	Sheet* dashboardSheet = nullptr;

	if (book->load(L"d:/new.xlsx")) {
		// File exists, check for specific sheets

		for (int i = 0; i < book->sheetCount(); ++i) {
			Sheet* sheet = book->getSheet(i);
			if (std::wcscmp(sheet->name(), L"project") == 0) {
				projectSheet = sheet;
				clearSheet(projectSheet);  // Assuming you have a clearSheet function defined
			}
			if (std::wcscmp(sheet->name(), L"dashboard") == 0) {
				dashboardSheet = sheet;
				clearSheet(dashboardSheet);  // Assuming you have a clearSheet function defined
			}
		}

		if (!projectSheet) {
			projectSheet = book->addSheet(L"project");
		}

		if (!dashboardSheet) {
			dashboardSheet = book->addSheet(L"dashboard");
		}
	}
	else {
		// File does not exist, create new file with sheets
		projectSheet = book->addSheet(L"project");  // Add and assign the 'project' sheet
		dashboardSheet = book->addSheet(L"dashboard");  // Add and assign the 'dashboard' sheet
	}

	write_project_header(book, projectSheet);

	for (int i = 0; i < m_totalProjectNum; i++) {
		write_project_body(book, projectSheet, &(m_pProjects[0][i]));  // Assuming write_project_body is defined
	}

	// Save and release
	book->save(L"d:/new.xlsx");
	book->release();
}

*/