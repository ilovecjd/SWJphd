#include "pch.h"
#include "GlobalEnv.h"
#include "Company.h"
#include "Creator.h"


CCompany::CCompany(GLOBAL_ENV* pEnv, CCreator* pCreator)
{	
	m_env = *pEnv;
	m_creator = *pCreator;
}

CCompany::~CCompany()
{
	ClearMemory();

}

void CCompany::ClearMemory()
{

}

// 초기화 : 전역환	경 변수, 프로젝트 갯수, 프로젝트 내용
//          테이블 초기화 (프로젝트, 인력, 비용, 수익)
BOOL CCompany::Init()
{   

	// 오더테이블의 기간은 넉넉히 2배로 설정해 놓는다.
	int maxTime = m_env.maxPeriod;

	m_doingHR.Resize(3, maxTime);
	m_totalHR.Resize(3, maxTime);
	m_freeHR.Resize(3, maxTime);
	m_doingTable.Resize(5, maxTime); // 크기는 자동으로 늘어나는거 맞지??
	m_incomeTable.Resize(1, maxTime);
	m_expensesTable.Resize(1, maxTime);
	m_balanceTable.Resize(1, maxTime);



	m_creator.m_pProjects;
	m_creator.m_totalProjectNum;




	// 매 시점 마다 필요한 비용
	int expense = m_env.higHrCount * m_env.higHrCost + m_env.midHrCount * m_env.midHrCost + m_env.lowHrCount * m_env.lowHrCost;

	m_doingHR.Clear(0);	// 일하는 사람도 없고
	for (int i = 0; i < maxTime; i++)
	{
		m_totalHR[HR_HIG][i] = m_freeHR[HR_HIG][i] = m_env.higHrCount;
		m_totalHR[HR_MID][i] = m_freeHR[HR_MID][i] = m_env.midHrCount;
		m_totalHR[HR_LOW][i] = m_freeHR[HR_LOW][i] = m_env.lowHrCount;
	}
	
	m_expensesTable.Clear(expense);// 고정비용 나가고
	m_doingTable.Clear(-1); // 하는일은 없고
	m_incomeTable.Clear(0); // 들어올돈 없고
	m_balanceTable.Clear(expense); // checkLastWeek 함수에서 자금 상황 체크용으로 사용함
	
    return TRUE;
}

// Order table 복구
void CCompany::ReadProject(FILE* fp)
{
	/*SAVE_TL tl;
	if (fread(&tl, 1, sizeof(tl), fp) != sizeof(tl)) {
		perror("Failed to read header");		
	}

	if (tl.type == TYPE_PROJECT) 
	{
		m_totalProjectNum = tl.length;
		m_AllProjects = new PROJECT[m_totalProjectNum];
	}

	fread(m_AllProjects, sizeof(PROJECT), m_totalProjectNum,  fp);*/
}


// 이번 기간에 결정할 일들. 프로젝트의 신규진행, 멈춤, 인원증감 결정
BOOL CCompany::Decision(int thisTime ) {

	m_lastDecisionTime = thisTime;

	//// 1. 지난주에 진행중인 프로젝트중 완료되지 않은 프로젝트들만 이번주로 이관
	if (FALSE == CheckLastWeek(thisTime))
	{
		//파산		
		return FALSE;
	}
	//
	//// 2. 진행 가능한 후보프로젝트들을  candidateTable에 모은다
	SelectCandidates(thisTime);

	//// 3. 신규 프로젝트 선택 및 진행프로젝트 업데이트
	//// 이번주에 발주된 프로젝트중 시작할 것이 있으면 진행 프로젝트로 기록
	SelectNewProject(thisTime);
	//
	////PrintDBData();
	return TRUE;
}


// 완료프로젝트 검사 및 진행프로젝트 업데이트
// 1. 지난 기간에 진행중인 프로젝트중 완료 내부 프로젝트가 있는가?
BOOL CCompany::CheckLastWeek(int thisTime)
{
	if (0 == thisTime) // 첫주는 체크할 지난주가 없음
		return TRUE;

	//지난주에 진행 중이던 프로젝트의 갯수
	int nLastProjects = m_doingTable[ORDER_SUM][thisTime - 1];

	// 외부는 무시, 내부는 종료되었으면 자금 테이블 갱신
	for (int i = 0; i < nLastProjects; i++)
	{
		int prjId = m_doingTable[i + 1][thisTime - 1];
		if (prjId == -1) // 지난주 진행 프로젝트 없음
			return TRUE;

		PROJECT* pProject = &(m_creator.m_pProjects[0][prjId]);

		if (pProject->category == INTERNAL_PRJ) // 내부프로젝트면
		{
			// 1. 지난주에 종료되었으면 앞으로 받을 금액표 업데이트
			if (pProject->endTime == (thisTime-1)) 	
			{
				for (int i = 0 ; i < m_env.lifeCycle; i++)
				{	
					m_incomeTable[0][thisTime + i] += pProject->revenue; // revenue 는 선택된 모드의 fixedIncome
				}			
			}
		}
	}
	
	// 자금 현황 체크. 전체를 매번 다시 계산하는것이 불편해 보여도 그대로 두자.
	// 현재 보유중인 현금
	int Cash = 0;// m_incomeTable의 첫 값을 m_GlobalEnv.Cash_Init; 로 초기화함.
	for (int i = 0; i < thisTime; i++)
	{
		Cash += (m_incomeTable[0][i] - m_expensesTable[0][i]);
		m_balanceTable[0][i] = Cash;
	}

	// 이번주 현금은 이상이 없는가?
	if (Cash < 0)// 이번주에 파산
	{
		return FALSE;
	}
	
	return TRUE;
}

void CCompany::SelectCandidates(int thisTime)
{	
	for (int i = 0; i< MAX_CANDIDATES; i++)
		m_candidateTable[i] = -1;

	int j = 0;
		
	// orderTable 에서 이번주 발생한 것들만 가져 올 수도 있지만 일단 편하게
	for (int i = 0; i < m_creator.m_totalProjectNum; i++)
	{
		PROJECT* pProject = &(m_creator.m_pProjects[0][i]);

		// 1. 이번 기간에 발생한것들중.
		// 2. 외부는 인원이 가능한가?
		// 3. 내부는 인원이 가능한 모드는 어느것인가?

		if(pProject->createTime == thisTime)
		{
			if ((pProject->category == EXTERNAL_PRJ))  // 외부 프로젝트
			{
				if (IsEnoughHR(thisTime, pProject)) // 인원 체크			
					m_candidateTable[j++] = pProject->ID;
			}
			else  //내부프로젝트
			{
				if (IsInternalEnoughHR(thisTime, pProject)) // 내부 프로젝트 인원 체크			
					m_candidateTable[j++] = pProject->ID;
			}
		}
	}
}

// 외부 프로젝트의 인원 체크
BOOL CCompany::IsEnoughHR(int thisTime, PROJECT* pProject)
{
	// 원본 인력 테이블을 복사해서 프로젝트 인력을 추가 할 수 있는지 확인한다.
	Dynamic2DArray doingHR = m_doingHR;
	int endTime = pProject->endTime;

	for (int i = thisTime; i <= endTime; i++)
	{
		doingHR[HR_HIG][i]	+= pProject->mode0.higHrCount;
		doingHR[HR_MID][i]	+= pProject->mode0.midHrCount;
		doingHR[HR_LOW][i]	+= pProject->mode0.lowHrCount;
	}

	for (int i = thisTime; i < m_env.maxPeriod; i++)
	{
		if (m_totalHR[HR_HIG][i] < doingHR[HR_HIG][i] ||
			m_totalHR[HR_MID][i] < doingHR[HR_MID][i] ||
			m_totalHR[HR_LOW][i] < doingHR[HR_LOW][i])
		{
			pProject->mode0.isPossible = FALSE;
			return FALSE;
		}
	}
	// 외부 프로젝트는 MODE0 만 사용한다.
	return TRUE;
}


// 내부 프로젝트를 위해 투입할 인원 체크
BOOL CCompany::IsInternalEnoughHR(int thisTime, PROJECT* pProject)
{
	// 원본 인력 테이블을 복사해서 프로젝트 인력을 추가 할 수 있는지 확인한다.
	Dynamic2DArray doingHR = m_doingHR;
	int endTime = pProject->endTime;

	for (int i = thisTime; i <= endTime; i++)
	{
		doingHR[HR_HIG][i] += pProject->mode0.higHrCount;
		doingHR[HR_MID][i] += pProject->mode0.midHrCount;
		doingHR[HR_LOW][i] += pProject->mode0.lowHrCount;
	}

	for (int i = thisTime; i < m_env.maxPeriod; i++)
	{
		if (m_totalHR[HR_HIG][i] < doingHR[HR_HIG][i] ||
			m_totalHR[HR_MID][i] < doingHR[HR_MID][i] ||
			m_totalHR[HR_LOW][i] < doingHR[HR_LOW][i])
		{
			pProject->mode0.isPossible = FALSE;
		}
	}

	doingHR = m_doingHR;

	for (int i = thisTime; i <= endTime; i++)
	{
		doingHR[HR_HIG][i] += pProject->mode1.higHrCount;
		doingHR[HR_MID][i] += pProject->mode1.midHrCount;
		doingHR[HR_LOW][i] += pProject->mode1.lowHrCount;
	}

	for (int i = thisTime; i < m_env.maxPeriod; i++)
	{
		if (m_totalHR[HR_HIG][i] < doingHR[HR_HIG][i] ||
			m_totalHR[HR_MID][i] < doingHR[HR_MID][i] ||
			m_totalHR[HR_LOW][i] < doingHR[HR_LOW][i])
		{
			pProject->mode1.isPossible = FALSE;
		}
	}

	doingHR = m_doingHR;

	for (int i = thisTime; i <= endTime; i++)
	{
		doingHR[HR_HIG][i] += pProject->mode2.higHrCount;
		doingHR[HR_MID][i] += pProject->mode2.midHrCount;
		doingHR[HR_LOW][i] += pProject->mode2.lowHrCount;
	}

	for (int i = thisTime; i < m_env.maxPeriod; i++)
	{
		if (m_totalHR[HR_HIG][i] < doingHR[HR_HIG][i] ||
			m_totalHR[HR_MID][i] < doingHR[HR_MID][i] ||
			m_totalHR[HR_LOW][i] < doingHR[HR_LOW][i])
		{
			pProject->mode2.isPossible = FALSE;
		}
	}

	if (pProject->mode0.isPossible == FALSE &&
		pProject->mode1.isPossible == FALSE &&
		pProject->mode2.isPossible == FALSE)
	{
		return FALSE;
	}
	
	return TRUE;// 하나의 모드라도 진행가능하면 TRUE
}


// 후보군들을 선택 정책에 따라서 순서를 변경한다.
// 2차원 배열을 오름차순으로 정렬하는 함수
void sortArrayAscending(int* indexArray, int* valueArray, int size) {
	// 두 배열을 정렬하기 위해 값과 인덱스를 페어로 묶어야 합니다.
	for (int i = 0; i < size - 1; i++) {
		for (int j = i + 1; j < size; j++) {
			if (valueArray[i] > valueArray[j]) {
				// 값(value)을 기준으로 정렬하고, 인덱스도 함께 변경합니다.
				std::swap(valueArray[i], valueArray[j]);
				std::swap(indexArray[i], indexArray[j]);
			}
		}
	}
}

// 2차원 배열을 내림차순으로 정렬하는 함수
void sortArrayDescending(int* indexArray, int* valueArray, int size) {
	for (int i = 0; i < size - 1; i++) {
		for (int j = i + 1; j < size; j++) {
			if (valueArray[i] < valueArray[j]) {
				// 값(value)을 기준으로 내림차순으로 정렬하고, 인덱스도 함께 변경합니다.
				std::swap(valueArray[i], valueArray[j]);
				std::swap(indexArray[i], indexArray[j]);
			}
		}
	}
}

void CCompany::SelectNewProject(int thisTime)
{	
	int valueArray[MAX_CANDIDATES] = {0, };  // 값 배열
	int j = 0;

	//while (m_candidateTable[j] != -1) {

	//	PROJECT* pProject;
	//	int id = m_candidateTable[j];

	//	pProject = &(m_creator.m_pProjects[0][id]);// m_AllProjects + id;

	//	if (0 < m_GlobalEnv.selectOrder) // 
	//	{
	//		if (pProject->category == 0) {// 외부 프로젝트
	//			valueArray[j] = pProject->profit;
	//		}
	//		else
	//		{
	//			int Order = m_GlobalEnv.selectOrder;
	//			if ((1 == Order) || ((4 == Order)))// 
	//				valueArray[j] = 0;// project->profit * 4 * 12 * 3;
	//			else
	//				valueArray[j] = 99999;// project->profit * 4 * 12 * 3;
	//		}
	//	}
	//	else // 0번 정책은 내부 외부 구분 없다.
	//	{
	//		valueArray[j] = pProject->profit;
	//	}

	//	j ++;
	//}
	//
	// 설정된 우선순위대로 프로젝트를 재 배치 한다.
	//switch (m_GlobalEnv.selectOrder)
	//{
	//case 0: // 발생 순서대로
	//	break;
	//case 1: // 내부를 먼저 외부는 작은것 위주 0
	//	sortArrayAscending(m_candidateTable, valueArray, j);	// 금액 오름차순 정렬	
	//	break;
	//case 2: // 내부를 마지막에 외부는 작은것 위주 99999
	//	sortArrayAscending(m_candidateTable, valueArray, j);	// 금액 오름차순 정렬	
	//	break;

	//case 3: // 내부를 먼저 외부는 큰것 위주 99999		
	//	sortArrayDescending(m_candidateTable, valueArray, j); // 금액 내림차순 정렬	
	//	break;
	//case 4: // 내부를 마지막에 외부는 큰것 위주 0
	//	sortArrayDescending(m_candidateTable, valueArray, j); // 금액 내림차순 정렬	
	//	break;

	//default : 
	//	break;
	//} 

	int i = 0;
	while (m_candidateTable[i] != -1) {

		if (i > MAX_CANDIDATES) break;

		int id = m_candidateTable[i++];

		PROJECT* pProject = &(m_creator.m_pProjects[0][id]);

		//if (0 < m_env.selectOrder)
		//{
			//if (pProject->category == 0)// 외부 프로젝트면
			//{
			//	if (pProject->startAbleTime < m_env.maxPeriod)
			//	{
			//		if (IsEnoughHR(thisTime, pProject))
			//		{
			//			AddProjectEntry(pProject, thisTime);
			//		}
			//	}
			//}
			//else  // 내부프로젝트면 
			//{
			//	if (IsInternalEnoughHR(thisTime, pProject))
			//	{					
			//		AddInternalProjectEntry(pProject, thisTime);
			//	}
			//	else {
			//		//if (IsIntenalEnoughtNextWeekHR(thisTime, pProject)) {// 내부 프로젝트는 전체 진행에 인원이 모자라도 다음주에만 진행 가능해도 인력 배치한다.
			//			AddInternalProjectEntry(pProject, thisTime);
			//		//}
			//	}
			//}
		//}
		//else // 0번 프로젝트는 내부외부 구분 없음
		//{			
			if (pProject->startAbleTime < m_env.maxPeriod)
			{
				if (IsEnoughHR(thisTime, pProject))
				{
					AddProjectEntry(pProject, thisTime);
				}
			}
		//}		
	}
}

// 모든 체크가 끝나고 호출된다. 
// 수주 프로젝트를 진행하기로 선택하면 호출된다.
// 단지 변수들만 셑팅하자.
void CCompany::AddProjectEntry(PROJECT* pProject,  int addTime)
{	
	int endTime = pProject->endTime;

	// 인력관리 테이블 업데이트
	for (int i = addTime; i <= endTime; i++)
	{
		// 인력관리 테이블 업데이트
		m_doingHR[HR_HIG][i] += pProject->mode0.higHrCount;
		m_doingHR[HR_MID][i] += pProject->mode0.midHrCount;
		m_doingHR[HR_LOW][i] += pProject->mode0.lowHrCount;

		m_freeHR[HR_HIG][i] = m_totalHR[HR_HIG][i] - m_doingHR[HR_HIG][i];
		m_freeHR[HR_MID][i] = m_totalHR[HR_MID][i] - m_doingHR[HR_MID][i];
		m_freeHR[HR_LOW][i] = m_totalHR[HR_LOW][i] - m_doingHR[HR_LOW][i];
				
		// 프로젝트 관리 현황판 업데이트
		// 프로젝트는 취소 되지 않는다. 그렇기 때문에 프로젝트 진행 기간동안의 모든 일정과 인원, 비용을 한번에 기록한다.
		int sum = m_doingTable[0][i];
		m_doingTable[sum + 1][i] = pProject->ID;
		m_doingTable[0][i] = sum + 1;
	}

	// 수입 테이블 업데이트
	int incomeDate;
	incomeDate = pProject->firstPayTime;	// 선금 지급일
	m_incomeTable[0][incomeDate] += pProject->firstPay;
	
	incomeDate = pProject->secondPayTime;	// 2차 지급일
	m_incomeTable[0][incomeDate] += pProject->secondPay;

	incomeDate = pProject->finalPayTime;	// 3차 지급일
	m_incomeTable[0][incomeDate] += pProject->finalPay;
}

//
//// 내부 프로젝트의 이번주 인원 배정
//void CCompany::AddInternalProjectEntry(PROJECT* project, int thisWeek)
//{
//	//project->runningWeeks => 현재가지 진행된 week 수. 0 이면 시작안한 상태
//	int runningWeeks = project->runningWeeks;
//	int startAvail = project->startAvail;
//
//	//song  !!!! int gap = thisWeek - startAvail  - runningWeeks; 값들간에 문제는 없는지 체크
//	// 밀린 시간 = 경과 시간 - 진행한 시간
//	// 경과 시간 = 현재 시각 - 시작 가능 시작
//	int gap = thisWeek - startAvail - runningWeeks;
//
//	// 예외를 위해서 메세지 박스를 띄운다. 디버깅용
//
//	// Activity들의 원본은 그대로 두고 복사해온 곳에 밀린만큼 모든 기간을 수정한다.
//	int numAct = project->numActivities;
//
//	PACTIVITY pActivity = new ACTIVITY[numAct];
//	for (int i = 0; i < numAct; i++)
//	{
//		pActivity[i] = project->activities[i];
//		pActivity[i].startDate += gap;
//		pActivity[i].endDate += gap;
//
//		int startDate = pActivity[i].startDate;
//		int endDate = pActivity[i].endDate;
//
//		for (int j = startDate; j <= endDate; j++)
//		{
//			if (j >= thisWeek)
//			{
//				m_doingHR[HR_HIG][j] += pActivity[i].highSkill;
//				m_doingHR[HR_MID][j] += pActivity[i].midSkill;
//				m_doingHR[HR_LOW][j] += pActivity[i].lowSkill;
//
//				m_freeHR[HR_HIG][j] = m_totalHR[HR_HIG][j] - m_doingHR[HR_HIG][j];
//				m_freeHR[HR_MID][j] = m_totalHR[HR_MID][j] - m_doingHR[HR_MID][j];
//				m_freeHR[HR_LOW][j] = m_totalHR[HR_LOW][j] - m_doingHR[HR_LOW][j];
//			}
//		}
//	}
//
//	// 현황판 업데이트
//	int sum = m_doingTable[0][thisWeek];
//	m_doingTable[sum + 1][thisWeek] = project->ID;
//	m_doingTable[0][thisWeek] = sum + 1;
//
//	delete[] pActivity;
//}
//

int CCompany::CalculateFinalResult() 
{
	int result = m_env.initialFunds;

	for (int i = 0; i < m_lastDecisionTime; i++)
	{
		result += (m_incomeTable[0][i]- m_expensesTable[0][i]);
	}
	return result;
}

// 기간내 총 매출
int CCompany::CalculateTotalInCome()
{
	int result = 0;

	for (int i = 0; i < m_lastDecisionTime; i++)
	{
		result += m_incomeTable[0][i];
	}

	return result;
}

//void CCompany::PrintCompanyResualt(CString strFileName, CString strOutSheetName)
//{
//	Book* book = xlCreateXMLBook();  // Use xlCreateBook() for xls format	
//
//	// 정품 키 값이 들어 있다. 공개하는 프로젝트에는 포함되어 있지 않다. 
//	// 정품 키가 없으면 읽기가 300 컬럼으로 제한된다.
//#ifdef INCLUDE_LIBXL_KET	
//	book->setKey(_LIBXL_NAME, _LIBXL_KEY);
//#endif	
//	
//	Sheet* resultSheet = nullptr;
//
//	if (book->load((LPCWSTR)strFileName)) {
//
//		for (int i = 0; i < book->sheetCount(); ++i) {
//			Sheet* sheet = book->getSheet(i);
//			if (std::wcscmp(sheet->name(), strOutSheetName) == 0) {
//				resultSheet = sheet;
//				clearSheet(resultSheet);  // Assuming you have a clearSheet function defined
//			}
//		}
//
//		if (!resultSheet) {
//			resultSheet = book->addSheet(strOutSheetName);
//		}
//	}
//	else {
//		// File does not exist, create new file with sheets
//		resultSheet = book->addSheet(strOutSheetName);
//	}
//
//	write_CompanyInfo(book, resultSheet);
//
//	// Save and release
//	book->save((LPCWSTR)strFileName);
//	book->release();
//}
//
//void CCompany::write_CompanyInfo(Book* book, Sheet* sheet)
//{
//	sheet->writeStr(0, 0, _T("종료주"));
//	sheet->writeNum(0, 1, m_lastDecisionWeek);
//
//	int index = 1;// 엑셀의 2행 부터 기록 => 행의 시작이 0번 인덱스
//	draw_outer_border(book, sheet, index, 0, index + 2, m_GlobalEnv.SimulationWeeks, BORDERSTYLE_THIN, COLOR_BLACK);
//	sheet->writeStr(index++, 0, _T("주"));
//	sheet->writeStr(index++, 0, _T("누계"));
//	sheet->writeStr(index++, 0, _T("발주"));	
//
//	for (int col = 0; col < m_GlobalEnv.SimulationWeeks; ++col) {
//		sheet->writeNum(1, col+1, col);
//		sheet->writeNum(2, col+1, m_orderTable[0][col]);
//		sheet->writeNum(3, col+1, m_orderTable[1][col]);
//
//		sheet->writeNum(20, col + 1, col);// 21행에도 몇주인지 적어 주자.
//	}
//
//	index = 5; // 엑셀의 6행 부터 기록	
//	draw_outer_border(book, sheet, index+1, 0, index + 3, m_GlobalEnv.SimulationWeeks, BORDERSTYLE_THIN, COLOR_BLACK);
//	sheet->writeStr(index++, 0, _T("투입"));
//	sheet->writeStr(index++, 0, _T("H"));
//	sheet->writeStr(index++, 0, _T("M"));
//	sheet->writeStr(index++, 0, _T("L"));	
//
//	index = 10; // 엑셀의 11행 부터 기록
//	draw_outer_border(book, sheet, index + 1, 0, index + 3, m_GlobalEnv.SimulationWeeks, BORDERSTYLE_THIN, COLOR_BLACK);
//	sheet->writeStr(index++, 0, _T("여유"));
//	sheet->writeStr(index++, 0, _T("H"));
//	sheet->writeStr(index++, 0, _T("M"));
//	sheet->writeStr(index++, 0, _T("L"));
//		
//	index = 15; // 엑셀의 16행 부터 기록
//	draw_outer_border(book, sheet, index + 1, 0, index + 3, m_GlobalEnv.SimulationWeeks, BORDERSTYLE_THIN, COLOR_BLACK);
//	sheet->writeStr(index++, 0, _T("총원"));
//	sheet->writeStr(index++, 0, _T("H"));
//	sheet->writeStr(index++, 0, _T("M"));
//	sheet->writeStr(index++, 0, _T("L"));	
//
//	for (int col = 0; col < m_GlobalEnv.SimulationWeeks; ++col) {
//
//		index = 6; // 엑셀의 7행 부터 기록 ("투입");
//		sheet->writeNum(index++, col + 1, m_doingHR[HR_HIG][col]);
//		sheet->writeNum(index++, col + 1, m_doingHR[HR_MID][col]);
//		sheet->writeNum(index++, col + 1, m_doingHR[HR_LOW][col]);
//
//		index = 11; // 엑셀의 12행 부터 기록 ("여유")
//		sheet->writeNum(index++, col + 1, m_freeHR[HR_HIG][col]);
//		sheet->writeNum(index++, col + 1, m_freeHR[HR_MID][col]);
//		sheet->writeNum(index++, col + 1, m_freeHR[HR_LOW][col]);
//
//		index = 16; // 엑셀의 17행 부터 기록 ("총원")
//		sheet->writeNum(index++, col + 1, m_totalHR[HR_HIG][col]);
//		sheet->writeNum(index++, col + 1, m_totalHR[HR_MID][col]);
//		sheet->writeNum(index++, col + 1, m_totalHR[HR_LOW][col]);
//	}
//
//	for (int col = 0; col < m_GlobalEnv.SimulationWeeks; ++col) {
//
//		draw_outer_border(book, sheet, 21, 0, 21 + MAX_DOING_TABLE_SIZE - 1, m_GlobalEnv.SimulationWeeks, BORDERSTYLE_THIN, COLOR_BLACK);
//
//		for (int i = 0; i < MAX_DOING_TABLE_SIZE; i++)// 
//		{
//			index = 21; 
//			sheet->writeNum(index+i, col + 1, m_doingTable[i][col]);		
//		}
//
//		index += (MAX_DOING_TABLE_SIZE + 1);
//		sheet->writeNum(index , col + 1, m_incomeTable[0][col]);
//
//		index += 1;
//		sheet->writeNum(index , col + 1, m_expensesTable[0][col]);
//
//		index += 1;
//		sheet->writeNum(index, col + 1, m_balanceTable[0][col]);				
//	}
//}