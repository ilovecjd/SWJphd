#pragma once

#include "globalenv.h"
#include "Creator.h"



class CCompany
{
public:
	CCompany(GLOBAL_ENV* pEnv, CCreator* pCreator);
	~CCompany();

	GLOBAL_ENV*	m_pEnv;
	CCreator*	m_pCreator;

	Dynamic2DArray m_totalHR;
	Dynamic2DArray m_doingHR;
	Dynamic2DArray m_freeHR;

	Dynamic2DArray m_doingTable;
	//Dynamic2DArray m_doneTable;
	//Dynamic2DArray m_defferTable;

	Dynamic2DArray m_incomeTable;
	Dynamic2DArray m_expensesTable;
	Dynamic2DArray m_balanceTable;

	BOOL Init();

	void ClearMemory();
		
	BOOL Decision(int thisWeek);
	int CalculateFinalResult();
	int CalculateTotalInCome();		
	//void PrintCompanyResualt(CString strFileName, CString strOutSheetName);
	

	int m_lastDecisionTime;
	
	
private:
	// 초기화 필요한 변수들
	//int m_totalProjectNum;

	//PROJECT* m_AllProjects = NULL;	

	

	
		
	BOOL CheckLastWeek(int thisWeek);
	void SelectCandidates(int thisWeek);
	BOOL IsEnoughHR_ActMode(int thisTime, PROJECT* pProject);
	//BOOL IsEnoughHR(int thisWeek, PROJECT* project);
	//BOOL IsInternalEnoughHR(int thisTime, PROJECT* pProject); // 내부 프로젝트 인원 체크	
	//BOOL IsIntenalEnoughtNextWeekHR(int thisWeek, PROJECT* project);// 이번주 내부 프로젝트 인원 체크	
	void SelectNewProject(int thisWeek);
	
	int m_candidateTable[MAX_CANDIDATES] = { 0, };
	void AddProjectEntry(PROJECT* project, int addWeek);	
	void RemoveInternalProjectEntry(PROJECT* project, int thisWeek);	
	void ReadProject(FILE* fp);

	void AddInternalProjectEntry(PROJECT* project, int addWeek);
};