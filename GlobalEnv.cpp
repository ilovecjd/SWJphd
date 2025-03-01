
#include "pch.h"
#include "GlobalEnv.h"


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