#pragma once

#include <iostream>
#include <vector>
#include <algorithm>


// 등급별 인건비
#define HI_HR_COST 50
#define MI_HR_COST 40
#define LO_HR_COST 30

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

	double P0; // Probability0
	double P1; // Probability1
	double P2; // Probability2
	double P3; // Probability3
	double P4; // Probability4
	double P5; // Probability5
};

struct GLOBAL_ENV {
	
	int		simulationPeriod;
	int		maxWeek; //최대 80주(18개월)간 진행되는 프로젝트를 시뮬레이션 마지막에 기록할 수도 있다.	
	int		higSkillStaffCount;
	int		midSkillStaffCount;
	int		lowSkillStaffCount;
	int		initialFunds;
	double	avgWeeklyProjects;
	int		higSkillCostRate;
	int		midSkillCostRate;
	int		lowSkillCostRate;
	double	dwExpenseRate;	// 비용계산에 사용되는 제경비 비율
	//double	profitRate;		// 프로젝트 총비용 계산에 사용되는 제경비 비율
	int		selectOrder;	// 선택 순서  1: 먼저 발생한 순서대로 2: 금액이 큰 순서대로 3: 금액이 작은 순서대로

	_MODE_ENV mEnv0; 
	_MODE_ENV mEnv1;
	_MODE_ENV mEnv2;
};

//////////////////////////////////////////////////////////////
// 내부프로젝트의 3가지 모드
struct _MODE {
	int labor_h;	// 투입되는 고급 인력수
	int labor_m;	// 투입되는 중급 인력수
	int labor_l;	// 투입되는 초급 인력수
	int expense;	// 모드별 소요 비용
	int lifeCycle;	// 수익이 발생하는 기간(라이프사이클)
	int success;	// 내부 프로젝트의 성공확율
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
