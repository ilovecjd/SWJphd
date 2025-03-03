#pragma once

#include <iostream>
#include <vector>
#include <algorithm>

#define MAX_CANDIDATES 50

// Order Tabe index for easy reference
enum OrderIndex {
	ORDER_SUM = 0,
	ORDER_ORD,
	ORDER_COUNT // Total number of OrderTable
};

// HR Tabe index for easy reference
enum HRIndex {
	HR_HIG = 0,
	HR_MID,
	HR_LOW,
	HR_COUNT // Total number of HR Table
};

int PoissonRandom(double lambda);

// project 에서 사용
#define RND_HR_H  20
#define RND_HR_M  70


// Sheet enumeration for easy reference
enum SheetName {
	WS_NUM_GENV = 0,
	WS_NUM_DASHBOARD,
	WS_NUM_PROJECT,
	WS_NUM_ACTIVITY_STRUCT,
	WS_NUM_DEBUG_INFO,
	WS_TOTAL_SHEET_COUNT // Total number of sheets
};

extern LPOLESTR gSheetNames[WS_TOTAL_SHEET_COUNT];// = { L"parameters", L"dashboard", L"project", L"activity_struct", L"debuginfo" };

// mode를 구성하는 환경 변수를 기록
struct _MODE_ENV {
	// 
	int R0; // revenue0
	int R1; // revenue1
	int R2; // revenue2
	int R3; // revenue3
	int R4; // revenue4
	int R5; // revenue5
	int R6; // revenue6
	int R7; // revenue7
	int R8; // revenue8
	int R9; // revenue9
	int R10; // revenue10
	int R11; // revenue11
	int R12; // revenue12

	double P0; // Probability0
	double P1; // Probability1
	double P2; // Probability2
	double P3; // Probability3
	double P4; // Probability4
	double P5; // Probability5
	double P6; // Probability6
	double P7; // Probability7
	double P8; // Probability8
	double P9; // Probability9
	double P10; // Probability10
	double P11; // Probability11
	double P12; // Probability12
};

struct GLOBAL_ENV {
	//// 전역적인 환경	
	int		simulationPeriod;
	int		maxPeriod; // 시뮬레이션 종료 이후의 값들도 추적하기 위해	

	//// 기본 환경
	int		higHrCount;		//최초에 보유한 고급 인력
	int		midHrCount;		//최초에 보유한 중급 인력
	int		lowHrCount;		//최초에 보유한 초급 인력

	int		higHrCost;		// 고급자 인건비(기간당)
	int		midHrCost;		// 중급자 인건비(기간당)
	int		lowHrCost;		// 초급자 인건비(기간당)

	int		initialFunds;	//최초에 보유한 자금 (6개월 유지비)

	double	extPrjInTime;	// 기간내 평균적으로 발생하는 외부프로젝트의 갯수
	double	intPrjInTime;	// 기간내 평균적으로 발생하는 내부프로젝트의 갯수
	int		minDuration;	// 외부 프로젝트 최소 기간
	int		maxDuration;	// 외부 프로젝트 최대 기간
	int		durRule;		// 외부 프로젝트 기간의 분포규칙 0:일양분포

	// 모드 생성 정보
	int		minMode;
	int		maxMode;
	int		lifeCycle;		// 제품의 라이프 사이클
	double	erevenueRate;	// 인건비대비 수익률(총수익을 구할때 사용)
	double	mu1Rate;		// mode0의 mu0 x mu1Rate = mode1의 mu1
	double	mu2Rate;		// mode1의 mu1 x mu2Rate = mode2의 mu2
	double	sigma1Rate;		// mode0의 sigma0 x sigma1Rate = mode1의 sigma1
	double	sigma2Rate;		// mode1의 sigma1 x sigma2Rate = mode2의 sigma2

	
	
	double	dwExpenseRate;	// 비용계산에 사용되는 제경비 비율
	//double	profitRate;		// 프로젝트 총비용 계산에 사용되는 제경비 비율
	int		selectOrder;	// 선택 순서  1: 먼저 발생한 순서대로 2: 금액이 큰 순서대로 3: 금액이 작은 순서대로

};

//////////////////////////////////////////////////////////////
// 내부프로젝트의 3가지 모드
struct _MODE {
	int labor_h;	// 투입되는 고급 인력수
	int labor_m;	// 투입되는 중급 인력수
	int labor_l;	// 투입되는 초급 인력수
	int expense;	// 모드별 소요 비용
	int lifeCycle;	// 수익이 발생하는 기간(라이프사이클)
	double mu;		// 수익의 평균
	double sigma;	// 수익의 표준편차
	//int success;	// 내부 프로젝트의 성공확율
	int revenue;	// 완료후 매달 발생하는 수익
};



////////////////////////////////////////////////////////////////////
// 프로젝트 속성
struct PROJECT {
	int category;		// 프로젝트 분류 (0: 외부 / 1: 내부)
	int ID;				// 프로젝트의 번호
	int createTime;		// 발주일
	int startAbleTime;	// 시작 가능일
	int duration;		// 프로젝트의 총 기간
	int startTime;		// 프로젝트의 시작일
	int endTime;		// 프로젝트 종료일

	int revenue;		// 프로젝트의 수주금액
	int firstPay;		// 선금 액수
	int secondPay;		// 2차 지급 액수
	int finalPay;		// 3차 지급 액수
	int firstPayTime;	// 선금 지급일
	int secondPayTime;	// 2차 지급일
	int finalPayTime;	// 3차 지급일

	_MODE	mode0;
	_MODE	mode1;
	_MODE	mode2;

	int isStart;		// 진행 여부 (0: 미진행, 나머지: 진행시작한 주)
	int experience;	// 경험 (0: 무경험 1: 유경험)
	int winProb;		// 성공 확률
	int nCashFlows;	// 비용 지급 횟수
	int profit;	// 총 기대 수익 (HR 종속)
};


class DynamicProjectArray {
private:
	std::vector<std::vector<PROJECT>> data;

public:
	DynamicProjectArray() {}

	DynamicProjectArray(const DynamicProjectArray& other) : data(other.data) {} // 복사 생성자

	DynamicProjectArray& operator=(const DynamicProjectArray& other) { // 할당 연산자
		if (this != &other) {
			data = other.data;
		}
		return *this;
	}

	class Proxy {
	private:
		DynamicProjectArray& array;
		int rowIndex;

	public:
		Proxy(DynamicProjectArray& arr, int index) : array(arr), rowIndex(index) {}

		PROJECT& operator[](int colIndex) {
			if (colIndex >= array.data[rowIndex].size()) {
				array.data[rowIndex].resize(colIndex + 1); // 필요한 만큼 열 확장
			}
			return array.data[rowIndex][colIndex];
		}
	};

	Proxy operator[](int rowIndex) {
		if (rowIndex >= data.size()) {
			data.resize(rowIndex + 1); // 필요한 만큼 행 확장
		}
		return Proxy(*this, rowIndex);
	}

	void Resize(int x, int y) {
		data.resize(x);
		for (int i = 0; i < x; ++i) {
			data[i].resize(y);
		}
	}

	int getRows() const {
		return (int)data.size();
	}

	int getCols() const {
		if (!data.empty()) return (int)data[0].size();
		return 0;
	}
};


class Dynamic2DArray {
private:
	std::vector<std::vector<int>> data;

public:
	Dynamic2DArray() {}

	Dynamic2DArray(const Dynamic2DArray& other) : data(other.data) {} // 복사 생성자

	Dynamic2DArray& operator=(const Dynamic2DArray& other) { // 할당 연산자
		if (this != &other) {
			data = other.data;
		}
		return *this;
	}

	class Proxy {
	private:
		Dynamic2DArray& array;
		int rowIndex;

	public:
		Proxy(Dynamic2DArray& arr, int index) : array(arr), rowIndex(index) {}

		int& operator[](int colIndex) {
			if (colIndex >= array.data[rowIndex].size()) {
				array.data[rowIndex].resize(colIndex + 1, 0);
			}
			return array.data[rowIndex][colIndex];
		}
	};

	Proxy operator[](int rowIndex) {
		if (rowIndex >= data.size()) {
			data.resize(rowIndex + 1);
		}
		return Proxy(*this, rowIndex);
	}

	void Resize(int x, int y) {
		data.resize(x);
		for (int i = 0; i < x; ++i) {
			data[i].resize(y, 0);
		}
	}

	void copyFromContinuousMemory(int* src, int rows, int cols) {
		Resize(rows, cols);
		for (int i = 0; i < rows; ++i) {
			for (int j = 0; j < cols; ++j) {
				data[i][j] = src[i * cols + j];
			}
		}
	}

	// Copy data to an external continuous memory block
	void copyToContinuousMemory(int* dest, int maxElements) {
		int index = 0;
		for (int i = 0; i < data.size() && index < maxElements; ++i) {
			for (int j = 0; j < data[i].size() && index < maxElements; ++j) {
				dest[index++] = data[i][j];
			}
		}
	}

	int getRows() const {
		return (int)data.size();
	}

	int getCols() const {
		if (!data.empty()) return (int)data[0].size();
		return 0;
	}
};

// 확률에 따라서 0 또는 1 생성
int ZeroOrOneByProb(int probability);

// 랜덤 숫자 생성 함수
int RandomBetween(int low, int high);

// 일반 정규분포의 역함수 p: (0,1) 사이의 확률값 mu: 평균, sigma: 표준편차
double InverseNormal(double p, double mu, double sigma);

/*
일자 	2025-03-03						
구분							
Global	SimulTerm	156	시뮬레이션의 기간(주) = 156/52=3년				
기본환경	Hr_Init_H	2	최초에 보유한 고급 인력				
	Hr_Init_M	2	최초에 보유한 중급 인력				
	Hr_Init_L	2	최초에 보유한 초급 인력				
	higHrCost	50	고급자 인건비 (기간당)				
	midSkillCost	40	중급자 인건비 (기간당)				
	lowSkillCost	30	초급자 인건비 (기간당)				
	initialFunds	1440	최초 보유한 자금  ==> 6개월 유지비				
프로젝트생성정보	extPrjInTime	1.25	기간내 발생하는 평균 발주 프로젝트 수(포아송분포로)				
	intPrjInTime	0.25	기간내 발생하는 평균 내부 프로젝트 수(포아송분포로)				
	minDuration	 1 	외부 프로젝트 최소 기간				
	maxDuration	 12 	외부 프로젝트 최소 기간				
	durRule	0 	외부 프로젝트 기간의 분포규칙 0:일양분포				
모드생성정보	minMode	3 	mode 의 최소 기간				
	maxMode	12 	mode 의 최대 기간				
	lifeCycle	 12 	제품의 라이프 사이클				
	revenueRate	 2 	인건비대비 수익률 (총수익을 구할때 사용)				
	mu1Rate	 2.250 	mode0의 mu0 x mu1Rate = mode1의 mu1
	mu2Rate	 2.110 	mode1의 mu1 x mu2Rate = mode2의 mu2
	sigma1Rate	 1.500 	mode0의 sigma0 x sigma1Rate = mode1의 sigma1
	sigma2Rate	 1.333 	mode1의 sigma1 x sigma2Rate = mode2의 sigma2				

*/
