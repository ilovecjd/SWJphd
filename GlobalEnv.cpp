
#include "pch.h"
#include "GlobalEnv.h"
#include "XLEzAutomation.h"
#include "Company.h"
#include "Creator.h"

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
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->revenue);			j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->firstPay);			j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->secondPay);			j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->finalPay);			j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->firstPayTime);		j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->secondPayTime);		j++;
    pSaveXl->SetCellValue(SheetNum, posY + 2, j, pProject->finalPayTime);		j++;

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

    for (int i = startId; i < endId ;i++) {
        pProject = &(pCreator->m_pProjects[0][i]);
        PrintOneProject(pSaveXl, WS_NUM_PROJECT, pProject);
    }
    
    //PrintOneDashBoard();
 
    return;
}


void PrintDashBoard(CXLEzAutomation* pSaveXl, CCompany* pCompany, int thisTime)
{

}

void PrintDBoardHeader(CXLEzAutomation* pSaveXl, int SheetNum)
{
    CString strTitle[] = {
            _T("주"),_T("누계"),_T("발주"),_T("1")
          /*  _T("투입"),_T("HR_H"),_T("HR_M"),_T("HR_L"),_T("1"),
            _T("여유"),_T("HR_H"),_T("HR_M"),_T("HR_L"),_T("1"),
            _T("총원"),_T("HR_H"),_T("HR_M"),_T("HR_L")*/
    };

    int nPrintWidth = sizeof(strTitle) / sizeof(strTitle[0]);
    pSaveXl->WriteArrayToRange(SheetNum, 2, 2, strTitle, 1, nPrintWidth);
}


void PrintGenv(CXLEzAutomation* pSaveXl,int sheet, CCreator* pCreator )
{
    GLOBAL_ENV* pEnv = &(pCreator->m_env);
    int row = 0;
    int col = 1;

    pSaveXl->SetCellValue(sheet, 3, 1, _T("* 문제에 사용할 데이터를 생성한다."));
    pSaveXl->SetCellValue(sheet, 4, 1, _T("* 생성하는 데이터를 위한 기초정보를 기록한다."));
    pSaveXl->SetCellValue(sheet, 5, 1, _T("* 일자"));
    pSaveXl->SetCellValue(sheet, 6, 1, _T("구분"));
    pSaveXl->SetCellValue(sheet, 7, 1, _T("Global"));
    pSaveXl->SetCellValue(sheet, 10,1, _T("기본환경"));
    pSaveXl->SetCellValue(sheet, 17,1, _T("프로젝트생성정보"));
    pSaveXl->SetCellValue(sheet, 22,1, _T("모드생성정보"));
    pSaveXl->SetCellValue(sheet, 5, 2, _T("2025-03-07"));

    row = 7;
    col = 2;
    
    pSaveXl->SetCellValue(sheet, row, col++, _T("simulationPeriod")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->simulationPeriod); pSaveXl->SetCellValue(sheet, row, col++, _T("시뮬레이션의 기간(12개월 x 5년 = 60 개월)")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("maxPeriod")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->maxPeriod); pSaveXl->SetCellValue(sheet, row, col++, _T("시뮬레이션 종료 이후의 값들도 추적하기 위해(simulationPeriod x 2)")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("technicalFee")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->technicalFee); pSaveXl->SetCellValue(sheet, row, col++, _T("인건비대비 기술료율(총수익을 구할때 사용)")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("higHrCount")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->higHrCount); pSaveXl->SetCellValue(sheet, row, col++, _T("최초에 보유한 고급 인력")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("midHrCount")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->midHrCount); pSaveXl->SetCellValue(sheet, row, col++, _T("최초에 보유한 중급 인력")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("lowHrCount")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->lowHrCount); pSaveXl->SetCellValue(sheet, row, col++, _T("최초에 보유한 초급 인력")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("higHrCost")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->higHrCost); pSaveXl->SetCellValue(sheet, row, col++, _T("고급자 인건비(기간당)")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("midHrCost")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->midHrCost); pSaveXl->SetCellValue(sheet, row, col++, _T("중급자 인건비(기간당)")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("lowHrCost")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->lowHrCost); pSaveXl->SetCellValue(sheet, row, col++, _T("초급자 인건비(기간당)")); row++; col = 2;
    pSaveXl->SetCellValue(sheet, row, col++, _T("initialFunds")); pSaveXl->SetCellValue(sheet, row, col++, pEnv->initialFunds); pSaveXl->SetCellValue(sheet, row, col++, _T("최초 보유한 자금 ==> 6개월 유지비")); row++; col = 2;
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
            _T("category"), _T("ID"), _T("발주일"), _T("시작가능"),_T("기간"), _T("시작"), _T("끝"),
            _T("수익금"),_T("선금"),_T("중도금"),_T("잔금"),_T("선금일"), _T("중도금일"),_T("잔금일"),
            _T("고급"),_T("중급"),_T("초급"),_T("lifeCycle"),_T("MU"),_T("SIGMA"),_T("고정수익"),
            _T("고급"),_T("중급"),_T("초급"),_T("lifeCycle"),_T("MU"),_T("SIGMA"),_T("고정수익"),
            _T("고급"),_T("중급"),_T("초급"),_T("lifeCycle"),_T("MU"),_T("SIGMA"),_T("고정수익")
    };

    int nPrintWidth = sizeof(strTitle) / sizeof(strTitle[0]);

    pSaveXl->WriteArrayToRange(SheetNum, 1, 1, (CString*)strTitle, 1, nPrintWidth);

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
