
#include "pch.h"
#include "GlobalEnv.h"

#include <cmath>
#include <stdexcept>

int PoissonRandom(double lambda) {
	int k = 0;
	double p = 1.0;
	double L = exp(-lambda);  // L = e^(-lambda)

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
    return mu + sigma * result;
}
