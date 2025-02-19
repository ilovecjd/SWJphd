﻿
// SWJphdDlg.cpp: 구현 파일
//

#include "pch.h"
#include "framework.h"
#include "SWJphd.h"
#include "SWJphdDlg.h"
#include "DlgProxy.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CSWJphdDlg 대화 상자


IMPLEMENT_DYNAMIC(CSWJphdDlg, CDialogEx);

CSWJphdDlg::CSWJphdDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_SWJPHD_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = nullptr;
}

CSWJphdDlg::~CSWJphdDlg()
{
	// 이 대화 상자에 대한 자동화 프록시가 있을 경우 이 대화 상자에 대한
	//  후방 포인터를 null로 설정하여
	//  대화 상자가 삭제되었음을 알 수 있게 합니다.
	if (m_pAutoProxy != nullptr)
		m_pAutoProxy->m_pDialog = nullptr;
}

void CSWJphdDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CSWJphdDlg, CDialogEx)
	ON_WM_CLOSE()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
END_MESSAGE_MAP()


// CSWJphdDlg 메시지 처리기

BOOL CSWJphdDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 이 대화 상자의 아이콘을 설정합니다.  응용 프로그램의 주 창이 대화 상자가 아닐 경우에는
	//  프레임워크가 이 작업을 자동으로 수행합니다.
	SetIcon(m_hIcon, TRUE);			// 큰 아이콘을 설정합니다.
	SetIcon(m_hIcon, FALSE);		// 작은 아이콘을 설정합니다.

	// TODO: 여기에 추가 초기화 작업을 추가합니다.

	return TRUE;  // 포커스를 컨트롤에 설정하지 않으면 TRUE를 반환합니다.
}

// 대화 상자에 최소화 단추를 추가할 경우 아이콘을 그리려면
//  아래 코드가 필요합니다.  문서/뷰 모델을 사용하는 MFC 애플리케이션의 경우에는
//  프레임워크에서 이 작업을 자동으로 수행합니다.

void CSWJphdDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 그리기를 위한 디바이스 컨텍스트입니다.

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 클라이언트 사각형에서 아이콘을 가운데에 맞춥니다.
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 아이콘을 그립니다.
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// 사용자가 최소화된 창을 끄는 동안에 커서가 표시되도록 시스템에서
//  이 함수를 호출합니다.
HCURSOR CSWJphdDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// 컨트롤러에서 해당 개체 중 하나를 계속 사용하고 있을 경우
//  사용자가 UI를 닫을 때 자동화 서버를 종료하면 안 됩니다.  이들
//  메시지 처리기는 프록시가 아직 사용 중인 경우 UI는 숨기지만,
//  UI가 표시되지 않아도 대화 상자는
//  남겨 둡니다.

void CSWJphdDlg::OnClose()
{
	if (CanExit())
		CDialogEx::OnClose();
}

void CSWJphdDlg::OnOK()
{
	if (CanExit())
		CDialogEx::OnOK();
}

void CSWJphdDlg::OnCancel()
{
	if (CanExit())
		CDialogEx::OnCancel();
}

BOOL CSWJphdDlg::CanExit()
{
	// 프록시 개체가 계속 남아 있으면 자동화 컨트롤러에서는
	//  이 애플리케이션을 계속 사용합니다.  대화 상자는 남겨 두지만
	//  해당 UI는 숨깁니다.
	if (m_pAutoProxy != nullptr)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}

	return TRUE;
}

