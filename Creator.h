#pragma once

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
	CString m_strEnvFilePath;
	int m_totalProjectNum;	
	Dynamic2DArray m_orderTable;
	
	BOOL Init(CString filePath, GLOBAL_ENV* pGlobalEnv);	
	void Save(CString filename, CString strInSheetName);
	int MakeMode(PROJECT* pProject, int duration, int intOrExt );
		
private:
	GLOBAL_ENV m_gEnv;
	DynamicProjectArray m_pProjects;

	int CreateOrderTable();
	int CreateProjects();
	BOOL CreateActivities(PROJECT* pProject);
	double CalculateHRAndProfit(PROJECT* pProject);
	double CalculateTotalLaborCost(int highCount, int midCount, int lowCount);
	double CalculateLaborCost(const std::string& grade);
	void CalculatePaymentSchedule(PROJECT* pProject);

	void WriteProjet(FILE* fp);

	int CreateAllProjects();
	int CreateInternalProject(int Id, int week);
	int CraterExternalProject(int Id, int week);
	BOOL CreateInternalActivities(PROJECT* pProject);

public:
	void PrintProjectInfo();
};