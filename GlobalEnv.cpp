﻿
#include "pch.h"
#include "GlobalEnv.h"
#include "XLEzAutomation.h"
#include "Company.h"
#include "Creator.h"

#include <vector>
#include <limits>
#include <cmath>

#include <stdexcept>


void PrintOneProject(CXLEzAutomation* pSaveXl, int SheetNum, PROJECT* pProject)
{
    int posY    = pProject->ID;
    int j       = 1;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->category);			j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->ID);				    j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->createTime);		    j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->startAbleTime);		j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->duration);			j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->startTime);			j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->endTime);		    j++;
    //pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->revenue);			j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->firstPay);			j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->secondPay);			j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->finalPay);			j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->firstPayTime);		j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->secondPayTime);		j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->finalPayTime);		j++;

    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->actMode.higHrCount);	j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->actMode.midHrCount);	j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->actMode.lowHrCount);	j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->actMode.lifeCycle); 	j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->actMode.mu);			j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->actMode.sigma);		j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->actMode.fixedIncome);	j++;

    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode0.higHrCount);	j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode0.midHrCount);	j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode0.lowHrCount);	j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode0.lifeCycle); 	j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode0.mu);			j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode0.sigma);		j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode0.fixedIncome);	j++;

    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode1.higHrCount);	j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode1.midHrCount);	j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode1.lowHrCount);	j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode1.lifeCycle); 	j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode1.mu);			j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode1.sigma);	    j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode1.fixedIncome);	j++;

    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode2.higHrCount);	j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode2.midHrCount);	j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode2.lowHrCount);	j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode2.lifeCycle);	j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode2.mu);			j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode2.sigma);		j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->mode2.fixedIncome);	j++;

    return;
}

void PrintAllProject(CXLEzAutomation* pSaveXl, CCreator* pCreator, int thisTime)
{



}



void PrintOneTime(CXLEzAutomation* pSaveXl, CCreator* pCreator, CCompany* pCompany, int thisTime)
{
 
    PROJECT* pProject = NULL;

    // 이 시점에서의 프로젝트 시작번호와 갯수
    int startId     = pCreator->m_orderTable[0][thisTime];
    int endId       = pCreator->m_orderTable[1][thisTime];
    endId           += startId;

    for (int i = startId; i < endId ;i++) {
        pProject = &(pCreator->m_pProjects[0][i]);
        PrintOneProject(pSaveXl, WS_NUM_PROJECT, pProject);
    }
    
    PrintDashBoard(pSaveXl, WS_NUM_DASHBOARD, pCompany, pCreator, thisTime);
 
    return;
}


void PrintDashBoard(CXLEzAutomation* pSaveXl,int sheet, CCompany* pCompany, CCreator* pCreator, int thisTime)
{
    int posY = 2;
    int posX = thisTime + 2;// 컬럼의 인덱스는 2 base로 시작한다.

    // 이번 기간까지 발생누계, 이번기간 발생    
    pSaveXl->SetCellValue(sheet, posY, posX, thisTime); posY++;
    pSaveXl->SetCellValue(sheet, posY, posX, pCreator->m_orderTable[0][thisTime]); posY++;
    pSaveXl->SetCellValue(sheet, posY, posX, pCreator->m_orderTable[1][thisTime]); posY++;

    // 총인원
    posY = 7;
    pSaveXl->SetCellValue(sheet, posY, posX, pCompany->m_totalHR[HR_HIG][thisTime]); posY++;
    pSaveXl->SetCellValue(sheet, posY, posX, pCompany->m_totalHR[HR_MID][thisTime]); posY++;
    pSaveXl->SetCellValue(sheet, posY, posX, pCompany->m_totalHR[HR_LOW][thisTime]); posY++;

    // 투입인원
    posY = 12;
    pSaveXl->SetCellValue(sheet, posY, posX, pCompany->m_doingHR[HR_HIG][thisTime]); posY++;
    pSaveXl->SetCellValue(sheet, posY, posX, pCompany->m_doingHR[HR_MID][thisTime]); posY++;
    pSaveXl->SetCellValue(sheet, posY, posX, pCompany->m_doingHR[HR_LOW][thisTime]); posY++;
    
    // 여유인원
    posY = 17;
    pSaveXl->SetCellValue(sheet, posY, posX, pCompany->m_freeHR[HR_HIG][thisTime]); posY++;
    pSaveXl->SetCellValue(sheet, posY, posX, pCompany->m_freeHR[HR_MID][thisTime]); posY++;
    pSaveXl->SetCellValue(sheet, posY, posX, pCompany->m_freeHR[HR_LOW][thisTime]); posY++;

    // 진행프로젝트/모드 형태로 표시한다.
    int doingProjects = pCompany->m_doingTable[ORDER_SUM][thisTime];

    posY = 21;
    pSaveXl->SetCellValue(sheet, posY, posX, doingProjects); posY++;

    for (int i = 0; i < doingProjects ; i++) // 
    {
        int ID              = pCompany->m_doingTable[i+1][thisTime];
        pSaveXl->SetCellValue(sheet, posY, posX,ID); posY++;
    }

    // 수입, 지출, 차액
    if (0 < thisTime)
    {
        posY = 28;
        pSaveXl->SetCellValue(sheet, posY, posX-1, pCompany->m_incomeTable[0] [thisTime-1] ); posY++;
        pSaveXl->SetCellValue(sheet, posY, posX-1, pCompany->m_expensesTable[0] [thisTime-1] ); posY++;    
        pSaveXl->SetCellValue(sheet, posY, posX-1, pCompany->m_balanceTable[0][thisTime-1]); posY++;
    }
}

void PrintDBoardHeader(CXLEzAutomation* pSaveXl, int SheetNum)
{
    CString strTitle[] = {
            _T("주"),_T("누계"),_T("발주"),_T(""),
            _T("총원"),_T("HR_H"),_T("HR_M"),_T("HR_L"),_T(""),
            _T("투입"),_T("HR_H"),_T("HR_M"),_T("HR_L"),_T(""),
            _T("여유"),_T("HR_H"),_T("HR_M"),_T("HR_L"),_T(""),
            _T("진행"),_T(""),_T(""),_T(""),_T(""),_T(""),_T(""),
            _T("수입"),_T("지출"),_T("차액"),
    };

    int nPrintWidth = sizeof(strTitle) / sizeof(strTitle[0]);
    pSaveXl->WriteArrayToRange(SheetNum, 2, 1, strTitle, nPrintWidth,1);
}


// is EvnFile : 환경변수 파일이면 true, 결과물 파일이면 fasle
void LoacalValueToExcel(CXLEzAutomation* pSaveXl,int sheet, GLOBAL_ENV* pEnv,BOOL isEnvFile )
{    
    int row = 0;
    int col = 1;

    SYSTEMTIME st;
    GetLocalTime(&st);  // 현재 로컬 날짜와 시간 얻기

    CString strDate, strTime;
    strDate.Format(_T("%04d-%02d-%02d"), st.wYear, st.wMonth, st.wDay);   // 예: "2025-03-07"
    strTime.Format(_T("%02d시%02d분%02d초"), st.wHour, st.wMinute, st.wSecond); // 예: "12시30분03초"

    //if (!isEnvFile)
    pSaveXl->SetCellValue(sheet, 3, 1, _T("* 문제에 사용할 데이터를 생성한다."));
    pSaveXl->SetCellValue(sheet, 4, 1, _T("* 생성하는 데이터를 위한 기초정보를 기록한다."));
    pSaveXl->SetCellValue(sheet, 5, 1, _T("* 일자"));
    pSaveXl->SetCellValue(sheet, 6, 1, _T("구분"));
    pSaveXl->SetCellValue(sheet, 7, 1, _T("Global"));
    pSaveXl->SetCellValue(sheet, 10,1, _T("기본환경"));
    pSaveXl->SetCellValue(sheet, 17,1, _T("프로젝트생성정보"));
    pSaveXl->SetCellValue(sheet, 22,1, _T("모드생성정보"));
    pSaveXl->SetCellValue(sheet, 5, 2, strDate);
    pSaveXl->SetCellValue(sheet, 5, 3, strTime);

    row = 7;
    col = 2;
    
    pSaveXl->SetCellValue(sheet, row, col++, _T("problemCount")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->problemCount); pSaveXl->SetCellValue(sheet, row, col++, _T("실험 횟수")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("selectedMode")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->selectedMode); pSaveXl->SetCellValue(sheet, row, col++, _T("모드 선택(Mode0 ~Mode3)")); row++; col = 2;    
    pSaveXl->SetCellValue(sheet, row, col++, _T("simulationPeriod")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->simulationPeriod); pSaveXl->SetCellValue(sheet, row, col++, _T("시뮬레이션의 기간(12개월 x 5년 = 60 개월)")); row++; col = 2;    
    pSaveXl->SetCellValue(sheet, row, col++, _T("maxPeriod")); 

    if (isEnvFile)  {   pSaveXl->SetCellFormula(sheet, row, col++, _T("=C9+24"));   }
    else            {   pSaveXl->SetCellValue(sheet, row, col++, pEnv->maxPeriod);  }

    pSaveXl->SetCellValue(sheet, row, col++, _T("시뮬레이션 종료 이후의 값들도 추적하기 위해(simulationPeriod x 2)")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("technicalFee")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->technicalFee); pSaveXl->SetCellValue(sheet, row, col++, _T("인건비대비 기술료율(총수익을 구할때 사용)")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("higHrCount")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->higHrCount); pSaveXl->SetCellValue(sheet, row, col++, _T("최초에 보유한 고급 인력")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("midHrCount")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->midHrCount); pSaveXl->SetCellValue(sheet, row, col++, _T("최초에 보유한 중급 인력")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("lowHrCount")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->lowHrCount); pSaveXl->SetCellValue(sheet, row, col++, _T("최초에 보유한 초급 인력")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("higHrCost")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->higHrCost); pSaveXl->SetCellValue(sheet, row, col++, _T("고급자 인건비(기간당)")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("midHrCost")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->midHrCost); pSaveXl->SetCellValue(sheet, row, col++, _T("중급자 인건비(기간당)")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("lowHrCost")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->lowHrCost); pSaveXl->SetCellValue(sheet, row, col++, _T("초급자 인건비(기간당)")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("fundsHoldTerm")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->fundsHoldTerm); pSaveXl->SetCellValue(sheet, row, col++, _T("자금 보유 기간 (몇개월간의 자본금을 가지고 시작할 것인가?)")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("initialFunds"));

    if(isEnvFile)   {   pSaveXl->SetCellFormula(sheet, row, col++, _T("=SUMPRODUCT(C12:C14,C15:C17)*C18"));    }
    else            {   pSaveXl->SetCellValue(sheet, row, col++, pEnv->initialFunds);                          }

    pSaveXl->SetCellValue(sheet, row, col++, _T("최초 보유한 자금  ==> fundHoldTerm 개월 유지비")); row++; col = 2;        
    pSaveXl->SetCellValue(sheet, row, col++, _T("extPrjInTime")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->extPrjInTime); pSaveXl->SetCellValue(sheet, row, col++, _T("기간내 발생하는 평균 발주 프로젝트 수(포아송분포로)")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("intPrjInTime")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->intPrjInTime); pSaveXl->SetCellValue(sheet, row, col++, _T("기간내 발생하는 평균 내부 프로젝트 수(포아송분포로)")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("minDuration")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->minDuration); pSaveXl->SetCellValue(sheet, row, col++, _T("외부 프로젝트 최소 기간")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("maxDuration")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->maxDuration); pSaveXl->SetCellValue(sheet, row, col++, _T("외부 프로젝트 최대 기간")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("durRule")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->durRule); pSaveXl->SetCellValue(sheet, row, col++, _T("외부 프로젝트 기간의 분포규칙 0:일양분포")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("minMode")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->minMode); pSaveXl->SetCellValue(sheet, row, col++, _T("mode 의 최소 기간")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("maxMode")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->maxMode); pSaveXl->SetCellValue(sheet, row, col++, _T("mode 의 최대 기간")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("lifeCycle")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->lifeCycle); pSaveXl->SetCellValue(sheet, row, col++, _T("제품의 라이프 사이클")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("profitRate")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->profitRate); pSaveXl->SetCellValue(sheet, row, col++, _T("mode0의 마진율(revenue 에 이값을 곱한다)")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("mu0Rate")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->mu0Rate); pSaveXl->SetCellValue(sheet, row, col++, _T("sigma0 : mu0 = mu0Rate : sigma0Rate")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("sigma0Rate")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->sigma0Rate); pSaveXl->SetCellValue(sheet, row, col++, _T("sigma0 : mu0 = 80 : 20")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("mu1Rate")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->mu1Rate); pSaveXl->SetCellValue(sheet, row, col++, _T("mode0의 mu x mu1Rate = mode1의 mu")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("sigma1Rate")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->sigma1Rate); pSaveXl->SetCellValue(sheet, row, col++, _T("mode0의 sigma x sigma1Rate = mode1의 sigma")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("mu2Rate")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->mu2Rate); pSaveXl->SetCellValue(sheet, row, col++, _T("mode1의 mu x mu2Rate = mode2의 mu")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("sigma2Rate")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->sigma2Rate); pSaveXl->SetCellValue(sheet, row, col++, _T("mode1의 sigma x sigma2Rate = mode2의 sigma")); row++; col = 2;
    
}

// 시트의 특정위치에 
void PrintProjectSheetHeader(CXLEzAutomation* pSaveXl, int SheetNum  )
{
//#define _PRINT_WIDTH 35	
    CString strTitle[] = {
            _T("category"), _T("ID"), _T("발주일"), _T("시작가능"),_T("기간"), _T("시작"), _T("끝"),//_T("수익금"),
            _T("선금"),_T("중도금"),_T("잔금"),_T("선금일"), _T("중도금일"),_T("잔금일"),
            _T("고급A"),_T("중급A"),_T("초급A"),_T("lifeCycleA"),_T("MUA"),_T("SIGMAA"),_T("고정수익"),
            _T("고급0"),_T("중급0"),_T("초급0"),_T("lifeCycle0"),_T("MU0"),_T("SIGMA0"),_T("고정수익"),
            _T("고급1"),_T("중급1"),_T("초급1"),_T("lifeCycle1"),_T("MU1"),_T("SIGMA1"),_T("고정수익"),
            _T("고급2"),_T("중급2"),_T("초급2"),_T("lifeCycle2"),_T("MU2"),_T("SIGMA2"),_T("고정수익")
    };

    int nPrintWidth = sizeof(strTitle) / sizeof(strTitle[0]);

    pSaveXl->WriteArrayToRange(SheetNum, 1, 1, (CString*)strTitle, 1, nPrintWidth);

}

void PrintResultHeader(CXLEzAutomation* pSaveXl, int sheet)
{    
    CString strTitle[] = {
            _T(""), _T("Mode0"), _T(""), _T("Mode1"), _T(""), _T("Mode2"), _T(""),
            _T("순번"),_T("기간"),_T("잔액"),_T("기간"),_T("잔액"), _T("기간"),_T("잔액"),
    };

    int nPrintWidth = sizeof(strTitle) / sizeof(strTitle[0]);

    pSaveXl->WriteArrayToRange(sheet, 1, 1, (CString*)strTitle, 2, nPrintWidth / 2);
}

int PoissonRandom(double lambda) {
	int k		= 0;
	double p	= 1.0;
	double L	= exp(-lambda);  // L = e^(-lambda)

	do {
		k++;
		p *= static_cast<double>(rand()) / (RAND_MAX + 1.0);
	} while (p > L);

	return k - 1;
}


// 확률에 따라서 0 또는 1 생성
int ZeroOrOneByProb(int probability)
{
	double randomProb = (double)rand() / RAND_MAX;
	return (randomProb <= (double)probability / 100) ? 1 : 0;
}

// 랜덤 숫자 생성 함수
int RandomBetween(int low, int high) {
	return low + rand() % (high - low + 1);
}


// 일반 정규분포의 역함수
// p: (0,1) 사이의 확률값
// mu: 평균, sigma: 표준편차
double InverseNormal(double p, double mu, double sigma)
{
    if (p <= 0.0 || p >= 1.0)
        throw std::invalid_argument("p는 0과 1 사이의 값이어야 합니다.");

    const double plow = 0.02425;
    const double phigh = 1 - plow;
    double q, r, result;

    if (p < plow)
    {
        // 하위 꼬리 영역: p < 0.02425
        q = std::sqrt(-2 * std::log(p));
        result = (((((0.0000430638 * q - 0.000258001) * q - 0.00129731) * q + 0.0038915) * q + 0.0103283) * q + 0.5) /
            ((((0.000143788 * q + 0.00108867) * q + 0.00549406) * q + 0.0175233) * q + 1);
    }
    else if (p <= phigh)
    {
        // 중앙 영역: 0.02425 <= p <= 0.97575
        q = p - 0.5;
        r = q * q;
        result = (((((-39.6968302866538 * r + 220.946098424521) * r - 275.928510446969) * r + 138.357751867269) * r - 30.6647980661472) * r + 2.50662827745924) * q /
            (((((-54.4760987982241 * r + 161.585836858041) * r - 155.698979859887) * r + 66.8013118877197) * r - 13.2806815528857) * r + 1);
    }
    else
    {
        // 상위 꼬리 영역: p > 0.97575
        q = std::sqrt(-2 * std::log(1 - p));
        result = -(((((0.0000430638 * q - 0.000258001) * q - 0.00129731) * q + 0.0038915) * q + 0.0103283) * q + 0.5) /
            ((((0.000143788 * q + 0.00108867) * q + 0.00549406) * q + 0.0175233) * q + 1);
    }

    // 일반 정규분포: mu + sigma * (표준 정규분포 역함수 값)
    double value = mu + sigma * result;

    // 리턴값이 0보다 작으면 0을 반환하도록 처리
    if (value < 0)
        value = 0;

    return value;
}




// MakeFreq 함수
/*
Dynamic2DArray* pArray : 값이 들어 있는 배열
int sourRow : pArray 에서 도수분포표를 만들들 데이터가 있는 첫 행, 0번 부터시작하는 인덱스임.
int sourCol : pArray 에서 도수분포표를 만들들 데이터가 있는 첫 열, 0번 부터시작하는 인덱스임.
int rowCnt : pArray 에서 도수분포표를 만들들 데이터가 있는 행의 갯수
int colCnt : pArray 에서 도수분포표를 만들들 데이터가 있는 열의 갯수
int desRow : 생성된 도수분포표를 적을 시작 행
int desCol : 생성된 도수분포표를 적을 시작 열
int classCnt : 계급 구간수
*/

BOOL MakeFreq(Dynamic2DArray* pArray,
    int sourRow, int sourCol,
    int rowCnt, int colCnt,
    int desRow, int desCol,
    int numClasses)
{
    if (!pArray || rowCnt <= 0 || colCnt <= 0 || numClasses <= 0)
        return FALSE;

    // 1. source 영역의 모든 데이터에서 전체 최소/최대값 구하기
    int globalMin = (std::numeric_limits<int>::max)();
    int globalMax = (std::numeric_limits<int>::min)();

    for (int i = sourRow; i < sourRow + rowCnt; i++) {
        for (int j = sourCol; j < sourCol + colCnt; j++) {
            int val = (*pArray)[i][j];
            if (val < globalMin)
                globalMin = val;
            if (val > globalMax)
                globalMax = val;
        }
    }

    // 2. 최소/최대값을 정수로 처리 (이미 정수이므로 floor/ceil은 생략)
    // 단, 모든 값이 동일할 경우 대비하여 범위를 1로 설정
    int range = globalMax - globalMin;
    if (range == 0)
        range = 1;

    // 3. 계급 폭(classWidth)을 올림하여 정수로 계산
    int classWidth = (range + numClasses - 1) / numClasses;
    if (classWidth == 0)
        classWidth = 1;

    // 4. 각 source 열별 도수분포 계산
    // freq[j][i] : source의 j번째 열에 대해 계급 i (0~numClasses-1)의 빈도수
    std::vector< std::vector<int> > freq(colCnt, std::vector<int>(numClasses, 0));
    for (int i = sourRow; i < sourRow + rowCnt; i++) {
        for (int j = 0; j < colCnt; j++) {
            int val = (*pArray)[i][sourCol + j];
            int classIndex = (val - globalMin) / classWidth;
            if (classIndex >= numClasses)
                classIndex = numClasses - 1;
            freq[j][classIndex]++;
        }
    }

    // 5. 각 source 열별 누적도수분포 계산
    std::vector< std::vector<int> > cumFreq(colCnt, std::vector<int>(numClasses, 0));
    for (int j = 0; j < colCnt; j++) {
        cumFreq[j][0] = freq[j][0];
        for (int i = 1; i < numClasses; i++) {
            cumFreq[j][i] = cumFreq[j][i - 1] + freq[j][i];
        }
    }

    // 6. 계산 결과를 pArray에 기록하기
    // 출력 레이아웃:
    // - desCol : 각 계급의 시작값 (globalMin + i*classWidth)
    // - desCol+1 ~ desCol+colCnt : 각 source 열의 도수
    // - desCol+1+colCnt ~ desCol+1+2*colCnt - 1 : 각 source 열의 누적도수
    for (int i = 0; i < numClasses; i++) {
        int binStart = globalMin + i * classWidth;
        (*pArray)[desRow + i][desCol] = binStart;
        // 도수분포 기록
        for (int j = 0; j < colCnt; j++) {
            (*pArray)[desRow + i][desCol + 1 + j] = freq[j][i];
        }
        // 누적도수분포 기록
        for (int j = 0; j < colCnt; j++) {
            (*pArray)[desRow + i][desCol + 1 + colCnt + j] = cumFreq[j][i];
        }
    }

    return TRUE;
}
