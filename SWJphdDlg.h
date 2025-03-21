﻿
// SWJphdDlg.h: 헤더 파일
//

#pragma once
#include "GlobalEnv.h"
#include "Creator.h"
#include"Company.h"
#include "XLEzAutomation.h"

class CSWJphdDlgAutoProxy;




// CSWJphdDlg 대화 상자
class CSWJphdDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CSWJphdDlg);
	friend class CSWJphdDlgAutoProxy;

// 생성입니다.
public:
	CSWJphdDlg(CWnd* pParent = nullptr);	// 표준 생성자입니다.
	virtual ~CSWJphdDlg();

// 대화 상자 데이터입니다.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_SWJPHD_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 지원입니다.


// 구현입니다.
protected:
	CSWJphdDlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// 생성된 메시지 맵 함수
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	DECLARE_MESSAGE_MAP()

	// song 새로 구현된 내용들
private:
	CString m_strEnvFilePath;  // 환경 구성 파일 경로 저장 변수
	CEdit m_editEnvFilePath;   // 에디트 컨트롤 변수

	CString m_strSaveFilePath;  // 저장 파일 경로 저장 변수
	CEdit m_editSaveFilePath;   // 에디트 컨트롤 변수
	
	void LoadFilePathFromRegistry(); // 레지스트리에서 경로 읽기
	void SaveFilePathToRegistry();   // 레지스트리에 경로 저장
	void PrintAllProject(CCreator* pCreator);	
	
	GLOBAL_ENV			m_gEnv; // 전역 환경변수들을 담는다
	CCreator*			m_pCreator;
	CXLEzAutomation*	m_pSaveXl;
	CCompany*			m_pCompany;
	int					m_stepByStepCnt;// 디버깅을 위해서 한타임씩 진행 할때 사용하는 컨틀롤변수

	void PrintAllProject(CCreator m_pCreator);
	void RunSimulator(_RESULT_* simResult);
	BOOL InitCompanyAndCreator();

	void ScreenValueToLocalValue();

public:
	afx_msg void OnBnClickedBtnEnvOpen();
	afx_msg void OnBnClickedBtnSaveFileName();
	afx_msg void OnBnClickedBtnEnvLoad();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnEnChangeEditMaxDuration();
	afx_msg void OnBnClickedBtnStepByStep();
	afx_msg void OnBnClickedBtnEnvSave();
	afx_msg void OnBnClickedBtnGo();
};
