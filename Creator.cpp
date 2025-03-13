#include "pch.h"
#include "GlobalEnv.h"
#include "Creator.h"
#include "XLEzAutomation.h"
//#include <cctype>   // toupper 함수를 사용하기 위해 필요

CCreator::CCreator()
{
}
CCreator::CCreator(int count)
{
}

CCreator::~CCreator()
{
}

BOOL CCreator::Init(GLOBAL_ENV* pGlobalEnv)
{
	*(&m_env) = *pGlobalEnv;    

    m_orderTable.Resize(2, m_env.maxPeriod);
	m_pProjects.Resize(0, 150);
    m_totalProjectNum = CreateAllProjects();

    return TRUE;
}

/*
1) 월간 내부 프로젝트 발생 확율과 외부 프로젝트 발생 확율을 다르게 한다.
2)
*/
int CCreator::CreateAllProjects()
{
    int projectId    = 0;
    int tempId   = 0;
	//CreateProjects(INTERNAL_PRJ, prjectId++, 0);// 내부는 하나 가지고 시작한다.

    for (int time   = 0; time < m_env.maxPeriod; time++)
    {
        int newCnt  = 0;
        
        newCnt = PoissonRandom(m_env.intPrjInTime);	// 이번기간에 발생하는 프로젝트 갯수
        for (int i = 0; i < newCnt; i++) // 내부 프로젝트 발생
        {
            CreateProjects(INTERNAL_PRJ, projectId++, time);
        }

        newCnt = PoissonRandom(m_env.extPrjInTime);	// 이번기간에 발생하는 프로젝트 갯수
        for (int i  = 0; i < newCnt; i++) // 외부 프로젝트 발생
        {
            CreateProjects(EXTERNAL_PRJ, projectId++, time);
        }

        m_orderTable[0][time]   = tempId;//누계
        m_orderTable[1][time]   = projectId - tempId;//발주
        tempId               = projectId;
    }

    return projectId;
}

int CCreator::CreateProjects(int category, int Id, int time)
{
    PROJECT Project;
    memset(&Project, -1, sizeof(struct PROJECT));

    // 내부 외부에 따라서 최소 기간과 최대 기간이 차이가 있음.
    int duration = 0;
    if (EXTERNAL_PRJ == category)
        duration = RandomBetween(m_env.minDuration, m_env.maxDuration);
    else
        duration = RandomBetween(m_env.minMode, m_env.maxMode);
    

    MakeModeAndRevenue(&Project, duration, category);		    // mode 를 만들고 모드별로 인원을 계산한다. 

    Project.category        = category;			    // 프로젝트 분류 (0: 외부 / 1: 내부)
    Project.ID              = Id;					// 프로젝트의 번호	

    Project.createTime      = time;					// 발주일
    Project.startAbleTime   = time;					// 시작 가능일. 내부는 바로 진행가능	
    Project.duration        = duration;				// 프로젝트의 총 기간
    Project.startTime       = -1;					// 프로젝트의 시작일
    Project.endTime         = time + duration - 1;	// 프로젝트 종료일

    int revenue             = Project.revenue;

    int pay1     = 0, pay2     = 0, pay3     = 0;    
    int payTime1 = 0, payTime2 = 0, payTime3 = 0;
    
    payTime1 = time;

    if (1 == duration)
    {
        pay1     = revenue;// 시작한 달에 전부 받음
    }
    else if (duration > 1 && duration <= 3)
    {
		pay1        = (int) (revenue * 0.5);
        pay2        = revenue - pay1;
        payTime2    = Project.endTime;// 선금 잔금만
    }
    else 
    {
        pay1        = (int) (revenue * 0.3);
        pay2        = pay1;
        pay3        = revenue - pay1 * 2;
        payTime2    = time + duration / 2;  // 중도금
        payTime3    = Project.endTime;      // 잔금
     }
    
    Project.firstPay        = pay1;
    Project.secondPay       = pay2;
    Project.finalPay        = pay3;
    Project.firstPayTime    = payTime1;
    Project.secondPayTime   = payTime2;
    Project.finalPayTime    = payTime3;

    m_pProjects[0][Id] = Project;// 여기가 맞는지, 함수 밖에서 하는게 맞는지 생각해 보자.

    return 0;
}

int CCreator::MakeModeAndRevenue(PROJECT* pProject, int duration, int category)
{
    //1. mode0의 인원 산정
    int nHigh = 0, nMid = 0, nLow = 0;

    // 기간이 1~3이면: 고급 또는 중급 1명 (둘 중 하나만 할당)
    if (duration >= 1 && duration <= 3)
    {
        if (rand() % 2 == 0)    nHigh = 1;
        else                    nMid = 1;
    }
    // 기간이 4~5이면: 고급 1명, 중급 1명 또는 2명 (무작위로 결정)
    else if (duration >= 4 && duration < 7)
    {
        nHigh = 1;
        nMid = (rand() % 2) + 1; // 1 또는 2
    }
    // 기간이 6~12이면: 고급 1명, 중급 1명 또는 2명, 초급 1명 또는 2명 (무작위로 결정)
    else if (duration >= 7 && duration <= 12)
    {
        nHigh = 1;
        nMid = (rand() % 2) + 1;    // 1 또는 2
        nLow = (rand() % 2) + 1; // 1 또는 2
    }
    else
    {
        AfxMessageBox(_T("기간은 1에서 12 사이의 값이어야 합니다."));
        return 0;
    }

    //2. 프로젝트 수익(인건비x기간x기술료비율), mode0의 인력 기록
    double expense      = (nHigh*m_env.higHrCost +  nMid*m_env.midHrCost + nLow*m_env.lowHrCost) * duration;    
    int revenue         = (int)(expense * m_env.technicalFee); // 전체 이익은 기술료 비율만큼 크게
    pProject->revenue   = revenue;
    
    _MODE tempMode; //mode 0 
    RtlSecureZeroMemory(&tempMode,sizeof(_MODE));
    tempMode.higHrCount = nHigh;
    tempMode.midHrCount = nMid;
    tempMode.lowHrCount = nLow;
    
    if (category == INTERNAL_PRJ) // 내부 프로젝트이면
    { 
        int lifeCycle, profitRate;
        double p, mu, sigma, interRevenue;
        profitRate				= (int) m_env.profitRate;
        lifeCycle				= m_env.lifeCycle;
		pProject->revenue		= 0; // 내부는 수익이 없다.

        //3. mode0를 위한 mu와 sigma, 수익 계산
        p						= static_cast<double>(rand()) / RAND_MAX;  // 0과 1 사이의 난수;
        mu						= revenue * profitRate;// 평균은 인건비 수준 x 이익율
        sigma					= mu * m_env.mu0Rate / m_env.sigma0Rate;
        interRevenue			= InverseNormal(p, mu, sigma); // mode0의 전체수익, 언제나 0 이상으로 나옴

        tempMode.lifeCycle		= lifeCycle;
        tempMode.mu				= mu;       		// 수익의 평균(내부프로젝트만 의미를 가짐)
        tempMode.sigma			= sigma;
        tempMode.fixedIncome	= (int)(interRevenue / lifeCycle); // lifeCycle 동안 한 time에 얻는 수익
        tempMode.isPossible     = TRUE;

        pProject->mode0 = tempMode;

		//4. mode1을 위한 mu와 sigma, 수익 계산
        tempMode.higHrCount     = nHigh * 2;
        tempMode.midHrCount     = nMid * 2;
        tempMode.lowHrCount     = nLow * 2;

        p                       = static_cast<double>(rand()) / RAND_MAX;  // 0과 1 사이의 난수
        mu                      = mu * m_env.mu1Rate;// mode1평균은 mode0의 mu1Rate수준
        sigma                   = mu * m_env.sigma1Rate; //mode1의 sigma는 mode의 sigma1Rate수준
        interRevenue            = InverseNormal(p, mu, sigma);

        tempMode.lifeCycle      = lifeCycle;
        tempMode.mu             = mu;
        tempMode.sigma          = sigma;
        tempMode.fixedIncome    = (int)(interRevenue / lifeCycle); // lifeCycle 동안 한 time에 얻는 수익
        tempMode.isPossible     = TRUE;

        pProject->mode1			= tempMode;

		//5. mode2을 위한 mu와 sigma, 수익 계산
        tempMode.higHrCount		= nHigh * 4;
        tempMode.midHrCount		= nMid * 4;
        tempMode.lowHrCount		= nLow * 4;

        p						= static_cast<double>(rand()) / RAND_MAX;  // 0과 1 사이의 난수
        mu						= mu * m_env.mu2Rate;// mode2평균은 mode1의 mu2Rate수준
        sigma					= mu * m_env.sigma2Rate; //mode2의 sigma는 mode1의 sigma2Rate수준
        interRevenue			= InverseNormal(p, mu, sigma);

        tempMode.lifeCycle		= lifeCycle;
        tempMode.mu				= mu;
        tempMode.sigma			= sigma;
        tempMode.fixedIncome	= (int)(interRevenue / lifeCycle); // lifeCycle 동안 한 time에 얻는 수익
        tempMode.isPossible     = TRUE;

        pProject->mode2			= tempMode;
    }
    else
        pProject->mode0 = tempMode;


    return 0;
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

void CCompany::PrintProjects(CXLEzAutomation* pXl)
{
    /////////////////////////////////////////////////////////////////////////
    // project 시트에 헤더 출력
    CString strTitle[2][16] = {
        {
            _T("Category"), _T("PRJ_ID"), _T("기간"), _T("시작가능"), _T("끝"),
            _T("발주일"), _T("총수익"), _T("경험"), _T("성공%"), _T("CF갯수"),
            _T("CF1%"), _T("CF2%"), _T("CF3%"), _T("선금"), _T("중도"), _T("잔금")
        },
        {
            _T("act갯수"), _T(""), _T("Dur"), _T("start"), _T("end"),
            _T(""), _T("HR_H"), _T("HR_M"), _T("HR_L"), _T(""),
            _T("mon_cf1"), _T("mon_cf2"), _T("mon_cf3"), _T(""), _T("prjType"), _T("actType")
        }
    };
    pXl->WriteArrayToRange(WS_NUM_PROJECT, 1, 1, (CString*)strTitle, 2, 16);
    pXl->SetRangeBorder(WS_NUM_PROJECT, 1, 1, 2, 16, 1, xlThin, RGB(0, 0, 0));


    for (int i = 0; i < m_totalProjectNum; i++)
    {
        PROJECT* pProject = m_AllProjects + i;
        PrintProjectInfo(pXl, pProject);
    }


}

void CCompany::PrintProjectInfo(CXLEzAutomation* pXl, PROJECT* pProject)
{
    const int iWidth = 16;
    const int iHeight = 7;
    int posX, posY;

    // VARIANT 배열 생성 하고 VT_EMPTY로 초기화
    VARIANT projectInfo[iHeight][iWidth];

    for (int i = 0; i < iHeight; ++i) {
        for (int j = 0; j < iWidth; ++j) {
            VariantInit(&projectInfo[i][j]);
            projectInfo[i][j].vt = VT_EMPTY;
        }
    }

    // 첫 번째 행 설정
    posX = 0; posY = 0;
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->category;
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->ID;
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->duration;
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->startAvail;
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->endDate;
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->orderDate;
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = static_cast<int>(pProject->profit);
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->experience;
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->winProb;

    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->nCashFlows;
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->cashFlows[0];
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->cashFlows[1];
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->cashFlows[2];

    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->firstPay;
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->secondPay;
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->finalPay;


    // 두 번째 행 설정
    posX = 0; posY = 1;
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->numActivities;

    posX = 10;  // 빈 칸을 건너뛰기
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->firstPayMonth;
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->secondPayMonth;
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->finalPayMonth;

    posX = 14;  // 빈 칸을 건너뛰기
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->projectType;
    projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->activityPattern;

    // 활동 데이터 설정
    for (int i = 0; i < pProject->numActivities; ++i) {
        // 인덱스를 문자열로 변환하고 "Activity" 접두사 추가
        CString strAct;
        strAct.Format(_T("Activity%02d"), i + 1);

        posX = 1; // 엑셀의 2행 2열부터 적는다.
        projectInfo[posY][posX].vt = VT_BSTR; projectInfo[posY][posX++].bstrVal = strAct.AllocSysString();
        projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->activities[i].duration;
        projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->activities[i].startDate;
        projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->activities[i].endDate;

        posX = 6;  // 두 열 건너뛰기
        projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->activities[i].highSkill;
        projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->activities[i].midSkill;
        projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->activities[i].lowSkill;

        posY++;
    }

    int printY = 4 + (pProject->ID - 1) * iHeight;
    pXl->WriteArrayToRange(WS_NUM_PROJECT, printY, 1, (VARIANT*)projectInfo, iHeight, iWidth);
    pXl->SetRangeBorderAround(WS_NUM_PROJECT, printY, 1, printY + iHeight - 1, iWidth + 1 - 1, 1, 2, RGB(0, 0, 0));
}
*/