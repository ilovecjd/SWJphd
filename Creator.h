﻿#pragma once

#include "GlobalEnv.h"
//#include <xlnt/xlnt.hpp>
#include <string>

class CCreator
{
public:
	CCreator();
	CCreator(int count);
	~CCreator();

	// song
public :
	//CString m_strEnvFilePath;
	int m_totalProjectNum;	
	Dynamic2DArray m_orderTable;	
    DynamicProjectArray m_pProjects;
	GLOBAL_ENV* m_pEnv;

	BOOL Init(GLOBAL_ENV* pGlobalEnv);
	int MakeModeAndRevenue(PROJECT* pProject, int duration, int intOrExt );
	void SetActMode(int iRule);
		
private:
	int CreateAllProjects();
	int CreateProjects(int category, int Id, int week);		
	double CalculateHRAndProfit(PROJECT* pProject);
	double CalculateTotalLaborCost(int highCount, int midCount, int lowCount);
	double CalculateLaborCost(const std::string& grade);
	void CalculatePaymentSchedule(PROJECT* pProject);	

public:
	void PrintProjectInfo();
};