
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




/////////////////////////////////////////////////////////////////////
// 파일 처리 루틴들
bool OpenFile(const CString& filename, const TCHAR* mode, FILE** fp) {

	// CString을 const wchar_t*로 변환 (유니코드 지원)
	const wchar_t* pFileName = filename.GetString();

	errno_t err = _wfopen_s(fp, pFileName, mode);
	if (err != 0 || *fp == nullptr) {
		perror("Failed to open file");
		return false;
	}
	return true;
}

void CloseFile(FILE** fp) {
	if (*fp != nullptr) {
		fclose(*fp);
		*fp = nullptr;
	}
}

ULONG WriteDataWithHeader(FILE* fp, int type, const void* data, size_t dataSize) {

	ULONG ulWritten = 0;
	ULONG ulTemp = 0;
	SAVE_TL tl = { type, static_cast<int>(dataSize) };
	ulTemp = fwrite(&tl, sizeof(tl), 1,fp);  // 먼저 데이터 타입 및 길이 정보를 쓴다
	ulWritten += ulTemp * sizeof(tl);

	ulTemp = fwrite(data, dataSize,1, fp);   // 실제 데이터 쓰기
	ulWritten += ulTemp * dataSize;

	return  ulWritten;
}

bool ReadDataWithHeader(FILE* fp, void* data, size_t expectedSize, int expectedType) {
	SAVE_TL tl;
	if (fread(&tl, 1, sizeof(tl), fp) != sizeof(tl)) {
		perror("Failed to read header");
		return false;
	}

	if (tl.type != expectedType || tl.length != expectedSize) {
		fprintf(stderr, "Data type or size mismatch\n");
		return false;
	}

	if (fread(data, 1, tl.length,fp) != tl.length) {
		perror("Failed to read data");
		return false;
	}

	return true;
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