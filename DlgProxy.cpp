
// DlgProxy.cpp: 구현 파일
//

#include "pch.h"
#include "framework.h"
#include "SWJphd.h"
#include "DlgProxy.h"
#include "SWJphdDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CSWJphdDlgAutoProxy

IMPLEMENT_DYNCREATE(CSWJphdDlgAutoProxy, CCmdTarget)

CSWJphdDlgAutoProxy::CSWJphdDlgAutoProxy()
{
	EnableAutomation();

	// 자동화 개체가 활성화되어 있는 동안 계속 애플리케이션을 실행하기 위해
	//	생성자에서 AfxOleLockApp를 호출합니다.
	AfxOleLockApp();

	// 애플리케이션의 주 창 포인터를 통해 대화 상자에 대한
	//  액세스를 가져옵니다.  프록시의 내부 포인터를 설정하여
	//  대화 상자를 가리키고 대화 상자의 후방 포인터를 이 프록시로
	//  설정합니다.
	ASSERT_VALID(AfxGetApp()->m_pMainWnd);
	if (AfxGetApp()->m_pMainWnd)
	{
		ASSERT_KINDOF(CSWJphdDlg, AfxGetApp()->m_pMainWnd);
		if (AfxGetApp()->m_pMainWnd->IsKindOf(RUNTIME_CLASS(CSWJphdDlg)))
		{
			m_pDialog = reinterpret_cast<CSWJphdDlg*>(AfxGetApp()->m_pMainWnd);
			m_pDialog->m_pAutoProxy = this;
		}
	}
}

CSWJphdDlgAutoProxy::~CSWJphdDlgAutoProxy()
{
	// 모든 개체가 OLE 자동화로 만들어졌을 때 애플리케이션을 종료하기 위해
	// 	소멸자가 AfxOleUnlockApp를 호출합니다.
	//  이러한 호출로 주 대화 상자가 삭제될 수 있습니다.
	if (m_pDialog != nullptr)
		m_pDialog->m_pAutoProxy = nullptr;
	AfxOleUnlockApp();
}

void CSWJphdDlgAutoProxy::OnFinalRelease()
{
	// 자동화 개체에 대한 마지막 참조가 해제되면
	// OnFinalRelease가 호출됩니다.  기본 클래스에서 자동으로 개체를 삭제합니다.
	// 기본 클래스를 호출하기 전에 개체에 필요한 추가 정리 작업을
	// 추가하세요.

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CSWJphdDlgAutoProxy, CCmdTarget)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CSWJphdDlgAutoProxy, CCmdTarget)
END_DISPATCH_MAP()

// 참고: IID_ISWJphd에 대한 지원을 추가하여
//  VBA에서 형식 안전 바인딩을 지원합니다.
//  이 IID는 .IDL 파일에 있는 dispinterface의 GUID와 일치해야 합니다.

// {c422e8a3-6872-4cc0-aecc-d4af0571f0a8}
static const IID IID_ISWJphd =
{0xc422e8a3,0x6872,0x4cc0,{0xae,0xcc,0xd4,0xaf,0x05,0x71,0xf0,0xa8}};

BEGIN_INTERFACE_MAP(CSWJphdDlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CSWJphdDlgAutoProxy, IID_ISWJphd, Dispatch)
END_INTERFACE_MAP()

// IMPLEMENT_OLECREATE2 매크로가 이 프로젝트의 pch.h에 정의됩니다.
// {e9c9df8b-8966-4de4-ae41-a2b2ad5c0b55}
IMPLEMENT_OLECREATE2(CSWJphdDlgAutoProxy, "SWJphd.Application", 0xe9c9df8b,0x8966,0x4de4,0xae,0x41,0xa2,0xb2,0xad,0x5c,0x0b,0x55)


// CSWJphdDlgAutoProxy 메시지 처리기
