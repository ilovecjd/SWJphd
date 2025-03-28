﻿#include "pch.h"
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
	m_pEnv = pGlobalEnv;

    m_orderTable.Resize(2, m_pEnv->maxPeriod);
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

	CreateProjects(INTERNAL_PRJ, projectId++, 0);// 내부는 하나 가지고 시작한다.

    for (int time   = 0; time < m_pEnv->maxPeriod; time++)
    {
        int newCnt  = 0;
        
        newCnt = PoissonRandom(m_pEnv->intPrjInTime);	// 이번기간에 발생하는 프로젝트 갯수
        for (int i = 0; i < newCnt; i++) // 내부 프로젝트 발생
        {
            CreateProjects(INTERNAL_PRJ, projectId++, time);
        }

        newCnt = PoissonRandom(m_pEnv->extPrjInTime);	// 이번기간에 발생하는 프로젝트 갯수
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
    memset(&Project, 0, sizeof(struct PROJECT));

    // 내부 외부에 따라서 최소 기간과 최대 기간이 차이가 있음.
    int duration = 0;
    if (EXTERNAL_PRJ == category)
        duration = RandomBetween(m_pEnv->minDuration, m_pEnv->maxDuration);
    else
        duration = RandomBetween(m_pEnv->minMode, m_pEnv->maxMode);
    
    MakeModeAndRevenue(&Project, duration, category);		    // mode 를 만들고 모드별로 인원을 계산한다. 
    
    Project.category        = category;			    // 프로젝트 분류 (0: 외부 / 1: 내부)
    Project.ID              = Id;					// 프로젝트의 번호	

    Project.createTime      = time;					// 발주일
    Project.startAbleTime   = time;					// 시작 가능일. 내부는 바로 진행가능	
    Project.duration        = duration;				// 프로젝트의 총 기간
    Project.tempDuration    = duration;				// 백업용
    Project.startTime       = -1;					// 프로젝트의 시작일
    Project.endTime         = time + duration - 1;	// 프로젝트 종료일
    Project.isStart        = -1;                   // 프로젝트 미시작

    if (EXTERNAL_PRJ == category) { // 외주프로젝트이면 
    
        int revenue             = Project.actMode.fixedIncome;

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
    }

    m_pProjects[0][Id] = Project;// 여기가 맞는지, 함수 밖에서 하는게 맞는지 생각해 보자.

    return 0;
}
/*
//MakeModeAndRevenue ver 3 : 인원수를 1,2,3 배로 하고, mode0의 확정 수익은 정규분포, mode1,2는 2배 3배로함.
int CCreator::MakeModeAndRevenue(PROJECT* pProject, int duration, int category)
{
    //0. mod1과 mod2의 배율
    int mode1rate = 2;
    int mode2rate = 3; // mode0 대비3배
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
    double expense = (nHigh * m_pEnv->higHrCost + nMid * m_pEnv->midHrCost + nLow * m_pEnv->lowHrCost) * duration;
    int revenue = (int)(expense * m_pEnv->technicalFee + expense); // 전체 이익은 기술료 비율만큼 크게
    pProject->revenue = revenue;

    _MODE tempMode; //mode 0 
    RtlSecureZeroMemory(&tempMode, sizeof(_MODE));
    tempMode.higHrCount = nHigh;
    tempMode.midHrCount = nMid;
    tempMode.lowHrCount = nLow;
    tempMode.fixedIncome = revenue; // 추후에 외부프로젝트의 경우 mode0 를 전부 복사하는 방법도 있을 수 있어서 표시만 해 놓는다.

    if (category == INTERNAL_PRJ) // 내부 프로젝트이면
    {
        int lifeCycle, profitRate;
        double p, mu, sigma, interRevenue;
        profitRate = (int)m_pEnv->profitRate;
        lifeCycle = m_pEnv->lifeCycle;
        pProject->revenue = 0; // 내부는 수익이 없다.

        //3. mode0를 위한 mu와 sigma, 수익 계산
        p = (static_cast<double>(rand()) + 1.0) / (RAND_MAX + 2.0);
        mu = revenue * profitRate;// 평균은 인건비 수준 x 이익율
        sigma = mu * m_pEnv->mu0Rate / m_pEnv->sigma0Rate;
        interRevenue = InverseNormal(p, mu, sigma); // mode0의 전체수익, 언제나 0 이상으로 나옴

        tempMode.lifeCycle = lifeCycle;
        tempMode.mu = mu;       		// 수익의 평균(내부프로젝트만 의미를 가짐)
        tempMode.sigma = sigma;
        tempMode.fixedIncome = (int)(interRevenue / lifeCycle); // lifeCycle 동안 한 time에 얻는 수익
        tempMode.isPossible = TRUE;

        pProject->mode0 = tempMode;
        pProject->actMode = tempMode;

        //4. mode1을 위한 mu와 sigma, 수익 계산
        tempMode.higHrCount = nHigh * mode1rate;
        tempMode.midHrCount = nMid * mode1rate;
        tempMode.lowHrCount = nLow * mode1rate;

        p = (static_cast<double>(rand()) + 1.0) / (RAND_MAX + 2.0);
        mu = mu * m_pEnv->mu1Rate;// mode1평균은 mode0의 mu1Rate수준
        sigma = mu * m_pEnv->sigma1Rate; //mode1의 sigma는 mode의 sigma1Rate수준
        //interRevenue = InverseNormal(p, mu, sigma);

        tempMode.lifeCycle = lifeCycle;
        tempMode.mu = mu;
        tempMode.sigma = sigma;
        //tempMode.fixedIncome = (int)(interRevenue / lifeCycle); // lifeCycle 동안 한 time에 얻는 수익
        tempMode.fixedIncome = (int)(interRevenue / lifeCycle) * mode1rate;
        tempMode.isPossible = TRUE;

        pProject->mode1 = tempMode;

        //5. mode2을 위한 mu와 sigma, 수익 계산
        tempMode.higHrCount = nHigh * mode2rate;
        tempMode.midHrCount = nMid * mode2rate;
        tempMode.lowHrCount = nLow * mode2rate;

        p = (static_cast<double>(rand()) + 1.0) / (RAND_MAX + 2.0);  // p는 (0, 1) 범위의 값을 갖게 됨
        mu = mu * m_pEnv->mu2Rate;// mode2평균은 mode1의 mu2Rate수준
        sigma = mu * m_pEnv->sigma2Rate; //mode2의 sigma는 mode1의 sigma2Rate수준
        //interRevenue = InverseNormal(p, mu, sigma);

        tempMode.lifeCycle = lifeCycle;
        tempMode.mu = mu;
        tempMode.sigma = sigma;
        //tempMode.fixedIncome = (int)(interRevenue / lifeCycle); // lifeCycle 동안 한 time에 얻는 수익
        tempMode.fixedIncome = tempMode.fixedIncome * mode2rate;
        tempMode.isPossible = TRUE;

        pProject->mode2 = tempMode;
    }
    else
    {
        pProject->mode0 = tempMode;
        pProject->mode1 = tempMode;
        pProject->mode2 = tempMode;
        pProject->actMode = tempMode;
    }


    return 0;
}
*/

//MakeModeAndRevenue ver 2 : 인원수를 1,2,3 배로 하고, 확정 수익은 정규분포로함.
/*
int CCreator::MakeModeAndRevenue(PROJECT* pProject, int duration, int category)
{
    //0. mod1과 mod2의 배율
    int mode1rate = 2;
    int mode2rate = 3; // mode0 대비3배
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
    double expense = (nHigh * m_pEnv->higHrCost + nMid * m_pEnv->midHrCost + nLow * m_pEnv->lowHrCost) * duration;
    int revenue = (int)(expense * m_pEnv->technicalFee + expense); // 전체 이익은 기술료 비율만큼 크게
    pProject->revenue = revenue;

    _MODE tempMode; //mode 0 
    RtlSecureZeroMemory(&tempMode, sizeof(_MODE));
    tempMode.higHrCount = nHigh;
    tempMode.midHrCount = nMid;
    tempMode.lowHrCount = nLow;
    tempMode.fixedIncome = revenue; // 추후에 외부프로젝트의 경우 mode0 를 전부 복사하는 방법도 있을 수 있어서 표시만 해 놓는다.

    if (category == INTERNAL_PRJ) // 내부 프로젝트이면
    {
        int lifeCycle, profitRate;
        double p, mu, sigma, interRevenue;
        profitRate = (int)m_pEnv->profitRate;
        lifeCycle = m_pEnv->lifeCycle;
        pProject->revenue = 0; // 내부는 수익이 없다.

        //3. mode0를 위한 mu와 sigma, 수익 계산
        p = (static_cast<double>(rand()) + 1.0) / (RAND_MAX + 2.0);
        mu = revenue * profitRate;// 평균은 인건비 수준 x 이익율
        sigma = mu * m_pEnv->mu0Rate / m_pEnv->sigma0Rate;
        interRevenue = InverseNormal(p, mu, sigma); // mode0의 전체수익, 언제나 0 이상으로 나옴

        tempMode.lifeCycle = lifeCycle;
        tempMode.mu = mu;       		// 수익의 평균(내부프로젝트만 의미를 가짐)
        tempMode.sigma = sigma;
        tempMode.fixedIncome = (int)(interRevenue / lifeCycle); // lifeCycle 동안 한 time에 얻는 수익
        tempMode.isPossible = TRUE;

        pProject->mode0 = tempMode;
        pProject->actMode = tempMode;

        //4. mode1을 위한 mu와 sigma, 수익 계산
        tempMode.higHrCount = nHigh * mode1rate;
        tempMode.midHrCount = nMid * mode1rate;
        tempMode.lowHrCount = nLow * mode1rate;

        p = (static_cast<double>(rand()) + 1.0) / (RAND_MAX + 2.0);
        mu = mu * m_pEnv->mu1Rate;// mode1평균은 mode0의 mu1Rate수준
        sigma = mu * m_pEnv->sigma1Rate; //mode1의 sigma는 mode의 sigma1Rate수준
        interRevenue = InverseNormal(p, mu, sigma);

        tempMode.lifeCycle = lifeCycle;
        tempMode.mu = mu;
        tempMode.sigma = sigma;
        tempMode.fixedIncome = (int)(interRevenue / lifeCycle); // lifeCycle 동안 한 time에 얻는 수익
        tempMode.isPossible = TRUE;

        pProject->mode1 = tempMode;

        //5. mode2을 위한 mu와 sigma, 수익 계산
        tempMode.higHrCount = nHigh * mode2rate;
        tempMode.midHrCount = nMid * mode2rate;
        tempMode.lowHrCount = nLow * mode2rate;

        p = (static_cast<double>(rand()) + 1.0) / (RAND_MAX + 2.0);  // p는 (0, 1) 범위의 값을 갖게 됨
        mu = mu * m_pEnv->mu2Rate;// mode2평균은 mode1의 mu2Rate수준
        sigma = mu * m_pEnv->sigma2Rate; //mode2의 sigma는 mode1의 sigma2Rate수준
        interRevenue = InverseNormal(p, mu, sigma);

        tempMode.lifeCycle = lifeCycle;
        tempMode.mu = mu;
        tempMode.sigma = sigma;
        tempMode.fixedIncome = (int)(interRevenue / lifeCycle); // lifeCycle 동안 한 time에 얻는 수익
        tempMode.isPossible = TRUE;

        pProject->mode2 = tempMode;
    }
    else
    {
        pProject->mode0 = tempMode;
        pProject->actMode = tempMode;
    }


    return 0;
}
*/


// mu와 sigma를 가지고 이익을 구하는 버전, 최초에 작성한 함수임.
int CCreator::MakeModeAndRevenue(PROJECT* pProject, int duration, int category)
{
    //0. mod1과 mod2의 배율
    int mode1rate = 2;
    int mode2rate = 4; // mode0 대비3배

    //1. mode0의 인원 산정
    int nHigh = 0, nMid = 0, nLow = 0;

    // 기간이 1~3이면: 고급 또는 중급 1명 (둘 중 하나만 할당)
    if (duration >= 1 && duration <= 3)
    {
        if (rand() % 1 == 0)    nHigh = 1;
        else                    nMid = 1;
    }
    // 기간이 4~5이면: 고급 1명, 중급 1명 또는 2명 (무작위로 결정)
    else if (duration >= 4 && duration < 7)
    {
        nHigh = 1;
        nMid = (rand() % 1) + 1; // 1 또는 2
    }
    // 기간이 6~12이면: 고급 1명, 중급 1명 또는 2명, 초급 1명 또는 2명 (무작위로 결정)
    else if (duration >= 7 && duration <= 12)
    {
        nHigh = 1;
        nMid = (rand() % 1) + 1;    // 1 또는 2
        nLow = (rand() % 1) + 1; // 1 또는 2
    }
    else
    {
        AfxMessageBox(_T("기간은 1에서 12 사이의 값이어야 합니다."));
        return 0;
    }

    //2. 프로젝트 수익(인건비x기간x기술료비율), mode0의 인력 기록
    double expense      = (nHigh* m_pEnv->higHrCost +  nMid* m_pEnv->midHrCost + nLow* m_pEnv->lowHrCost) * duration;
    int revenue         = (int)(expense * m_pEnv->technicalFee + expense); // 전체 이익은 기술료 비율만큼 크게
    //pProject->revenue   = revenue;
    
    _MODE tempMode; //mode 0 
    RtlSecureZeroMemory(&tempMode,sizeof(_MODE));
    tempMode.higHrCount     = nHigh;
    tempMode.midHrCount     = nMid;
    tempMode.lowHrCount     = nLow;
    tempMode.fixedIncome    = revenue; // 외부프로젝트의 전체수익 또는 내부프로젝트의 매달 받는 금액
    
    if (category == INTERNAL_PRJ) // 내부 프로젝트이면
    { 
        int lifeCycle, profitRate;
        double p, mu, sigma, interRevenue;
        profitRate				= (int)m_pEnv->profitRate;
        lifeCycle				= m_pEnv->lifeCycle;
		//pProject->revenue		= 0; // 내부는 수익이 없다.

        //3. mode0를 위한 mu와 sigma, 수익 계산
        p						= (static_cast<double>(rand()) + 1.0) / (RAND_MAX + 2.0);
        mu						= revenue * profitRate;// 평균은 인건비 수준 x 이익율
        sigma					= mu *  m_pEnv->sigma0Rate / m_pEnv->mu0Rate;
        interRevenue			= InverseNormal(p, mu, sigma); // mode0의 전체수익, 언제나 0 이상으로 나옴

        tempMode.lifeCycle		= lifeCycle;
        tempMode.mu				= mu;       		// 수익의 평균(내부프로젝트만 의미를 가짐)
        tempMode.sigma			= sigma;
        tempMode.fixedIncome	= (int)(interRevenue / lifeCycle); // lifeCycle 동안 한 time에 얻는 수익
        tempMode.isPossible     = TRUE;

        pProject->mode0         = tempMode;
        pProject->actMode       = tempMode;

		//4. mode1을 위한 mu와 sigma, 수익 계산
        tempMode.higHrCount     = nHigh * mode1rate;
        tempMode.midHrCount     = nMid * mode1rate;
        tempMode.lowHrCount     = nLow * mode1rate;

        p                       = (static_cast<double>(rand()) + 1.0) / (RAND_MAX + 2.0);
        mu                      = mu * m_pEnv->mu1Rate;// mode1평균은 mode0의 mu1Rate수준
        sigma                   = sigma * m_pEnv->sigma1Rate; //mode1의 sigma는 mode의 sigma1Rate수준
        interRevenue            = InverseNormal(p, mu, sigma);

        tempMode.lifeCycle      = lifeCycle;
        tempMode.mu             = mu;
        tempMode.sigma          = sigma;
        tempMode.fixedIncome    = (int)(interRevenue / lifeCycle); // lifeCycle 동안 한 time에 얻는 수익
        tempMode.isPossible     = TRUE;

        pProject->mode1			= tempMode;

		//5. mode2을 위한 mu와 sigma, 수익 계산
        tempMode.higHrCount		= nHigh * mode2rate;
        tempMode.midHrCount		= nMid * mode2rate;
        tempMode.lowHrCount		= nLow * mode2rate;

        p						= (static_cast<double>(rand()) + 1.0) / (RAND_MAX + 2.0);  // p는 (0, 1) 범위의 값을 갖게 됨
        mu						= mu * m_pEnv->mu2Rate;// mode2평균은 mode1의 mu2Rate수준
        sigma					= sigma * m_pEnv->sigma2Rate; //mode2의 sigma는 mode1의 sigma2Rate수준
        interRevenue			= InverseNormal(p, mu, sigma);

        tempMode.lifeCycle		= lifeCycle;
        tempMode.mu				= mu;
        tempMode.sigma			= sigma;
        tempMode.fixedIncome	= (int)(interRevenue / lifeCycle); // lifeCycle 동안 한 time에 얻는 수익
        tempMode.isPossible     = TRUE;

        pProject->mode2			= tempMode;
    }
    else
    { 
        pProject->mode0         = tempMode;
        pProject->mode1         = tempMode;
        pProject->mode2         = tempMode;
        pProject->actMode       = tempMode;
    }

    return 0;
}


// 내부프로젝트의 actMode를 주어진 룰대로 변경한다. 
// SetActMode ver2 : 기간을 2배,4배 늘리는 버전. 이것은 인원이 변동없는 버전과 함께 사용해야 한다.
/*
void CCreator::SetActMode(int iRule)
{
    int i = 0;
    for(i = 0; i<m_totalProjectNum; i++)
    {
        PROJECT* pProject = &(m_pProjects[0][i]);
        if (INTERNAL_PRJ == pProject->category)
        {   
            int tempDur = pProject->tempDuration;
            switch(iRule)
            {
                case 0: 
                    pProject->actMode = pProject->mode0;
                    pProject->duration = tempDur;
                    pProject->endTime = pProject->createTime + tempDur - 1;

                    break;
                case 1:
                    pProject->actMode = pProject->mode1;
                    pProject->duration = tempDur * 2;
                    pProject->endTime = pProject->createTime + tempDur*2 - 1;
                    break;
                case 2:
                    pProject->actMode = pProject->mode2;
                    pProject->duration = tempDur * 4;
                    pProject->endTime = pProject->createTime + tempDur * 4 - 1;
                    break;                    
                default : 
                    pProject->actMode = pProject->mode0;                   
                    pProject->duration = tempDur ;
                    pProject->endTime = pProject->createTime + tempDur - 1;
                    break;
            }
        }
    }
}
*/

void CCreator::SetActMode(int iRule)
{
    int i = 0;
    for (i = 0; i < m_totalProjectNum; i++)
    {
        PROJECT* pProject = &(m_pProjects[0][i]);
        if (INTERNAL_PRJ == pProject->category)
        {   
            switch (iRule)
            {
            case 0:
                pProject->actMode = pProject->mode0;
                break;
            case 1:
                pProject->actMode = pProject->mode1;
                break;
            case 2:
                pProject->actMode = pProject->mode2;
                break;
            default:
                pProject->actMode = pProject->mode0;
                break;
            }
        }
    }
}