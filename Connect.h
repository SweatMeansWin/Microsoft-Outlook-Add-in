#pragma once
#include "resource.h"

#include INTERFACE_HEADER
#include "Serializator.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;

// CConnect
class CConnect;
typedef public IDispEventSimpleImpl<1, CConnect, &__uuidof(Outlook::ApplicationEvents_11)> ApplicationEventSink;

class ATL_NO_VTABLE CConnect :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CConnect, &CLSID_Connect>,
	public IDispatchImpl<IConnect, &IID_IConnect, &LIBID_OutlookAddinLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &LIBID_AddInDesignerObjects, /* wMajor = */ 1>,
	public IDispatchImpl<Outlook::ApplicationEvents_11, &__uuidof(Outlook::ApplicationEvents_11), &__uuidof(Outlook::__Outlook), /* wMajor = */ 9, /* wMinor = */ 4>,
	ApplicationEventSink
{
public:
	CConnect();
	~CConnect();

	DECLARE_REGISTRY_RESOURCEID(IDR_CONNECT)

	BEGIN_COM_MAP(CConnect)
		COM_INTERFACE_ENTRY(IConnect)
		COM_INTERFACE_ENTRY2(IDispatch, _IDTExtensibility2)
		COM_INTERFACE_ENTRY(_IDTExtensibility2)
		COM_INTERFACE_ENTRY(Outlook::ApplicationEvents_11)
	END_COM_MAP()

	BEGIN_SINK_MAP(CConnect)
		SINK_ENTRY_INFO(1, __uuidof(ApplicationEvents_11), 0xf002, OnItemSend, &OnItemSendInfo)
		SINK_ENTRY_INFO(1, __uuidof(Outlook::ApplicationEvents_11), 0xf003, OnNewMail, &OnNewMailInfo)
	END_SINK_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()
	
	// _IDTExtensibility2 Methods
	STDMETHOD(OnConnection)(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY * * custom);
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom);
	//Visual studio events
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom) { return S_OK; }
	STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom) { return S_OK; }
	STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom) { return S_OK; }

private:
	Outlook::_ApplicationPtr m_pApplication;
	std::condition_variable m_threadNotifierCv;
	std::condition_variable m_workerFinishedCv;
	std::atomic_bool m_workerStopped;
	std::atomic_bool m_hasNewMails;
	std::mutex m_mutex;

	bool startWorker();
	void stopWorker();
	void workerFunction(const CConnect* pConnectInstance);

	STDMETHOD(OnNewMail)();
	STDMETHOD(OnItemSend)(LPDISPATCH Item, VARIANT_BOOL * Cancel);

	static _ATL_FUNC_INFO OnItemSendInfo;
	static _ATL_FUNC_INFO OnNewMailInfo;
};

OBJECT_ENTRY_AUTO(__uuidof(ConnectClass), CConnect)
