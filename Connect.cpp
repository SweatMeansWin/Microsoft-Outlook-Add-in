// Connect.cpp : Implementation of CConnect

#include "PreCompiled.h"
#include "Connect.h"
#include "Utils.h"
#include "UserFoldersStorage.h"
#include "OutlookInstanceStorage.h"
#include "OutlookMailProcesser.h"

const int EVENT_WAIT_TIME = 2000;
const DWORD THREAD_EX_WAIT_TIME = 5000;
const int MSG_TIMEOUT = 50;
LPCWSTR IPM_NOTE_MSG_CLASS = L"IPM.Note";
_ATL_FUNC_INFO CConnect::OnItemSendInfo = { CC_STDCALL, VT_EMPTY, 2,{ VT_DISPATCH, VT_BYREF | VT_BOOL } };
_ATL_FUNC_INFO CConnect::OnNewMailInfo = { CC_STDCALL, VT_EMPTY, 0 };

std::atomic_bool g_inited = false;
CSerializator g_serializator;

CConnect::CConnect()
	: m_pApplication(nullptr), m_hasNewMails(false), m_workerStopped(false)
{
	CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
}

CConnect::~CConnect()
{
	CoUninitialize();
}

STDMETHODIMP CConnect::OnConnection(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY ** custom) {
	std::lock_guard<std::mutex> lck(m_mutex);
	if (g_inited)
		return S_OK;

	g_inited = true;
	m_pApplication = Outlook::_ApplicationPtr(Application);
	ApplicationEventSink::DispEventAdvise(static_cast<IDispatch*>(m_pApplication), &__uuidof(Outlook::ApplicationEvents_11));
	g_serializator.Serialize(m_pApplication);
	startWorker();
	return S_OK;
}

STDMETHODIMP CConnect::OnDisconnection(ext_DisconnectMode RemoveMode, SAFEARRAY ** custom) {
	ApplicationEventSink::DispEventUnadvise(static_cast<IDispatch*>(m_pApplication), &__uuidof(Outlook::ApplicationEvents_11));
	stopWorker();
	return S_OK;
}

bool CConnect::startWorker()
{
	SECURITY_ATTRIBUTES sa = { 0 };
	sa.nLength = sizeof(sa);
	sa.lpSecurityDescriptor = nullptr;
	sa.bInheritHandle = TRUE;

	std::thread worker(&CConnect::workerFunction, this, this);
	worker.detach();

	return true;
}

void CConnect::stopWorker()
{
	m_workerStopped = true;
	m_threadNotifierCv.notify_all();
	// wait for thread end
	std::unique_lock<std::mutex> lck(m_mutex);
	m_workerFinishedCv.wait(lck);
}

void CConnect::workerFunction(const CConnect* pConnectInstance)
{
	//Deserialise application and init base objects
	auto pApplication = g_serializator.DeSerialize();
	if (!pApplication) {
		return;
	}

	COutlookInstanceStorage outlookStorage(pApplication);
	if (!outlookStorage.IsCreated()) {
		return;
	}
	CUserFoldersStorage foldersStorage(outlookStorage);

	while (!pConnectInstance->m_workerStopped) {
		//Check new user folders
		auto pNameSpaceFoldersCollection = outlookStorage.GetNameSpaceFolders();
		if (pNameSpaceFoldersCollection) {
			Outlook::MAPIFolderPtr pNameSpaceFolder = nullptr;
			HRESULT hr = pNameSpaceFoldersCollection->GetFirst(&pNameSpaceFolder);
			if (hr == S_OK && pNameSpaceFolder) {
				do
				{
					if (COutlookInstanceStorage::CheckMailFolder(pNameSpaceFolder) && !foldersStorage.Contains(pNameSpaceFolder)) {
						foldersStorage.Add(pNameSpaceFolder);
					}
				} while (pNameSpaceFoldersCollection->GetNext(&pNameSpaceFolder) == S_OK && pNameSpaceFolder);
			}
		}
		
		for (size_t idx = 0; idx < foldersStorage.Size(); ++idx) {
			// Get idx's folder and process it
			auto pNameSpaceFolder = foldersStorage[idx];
			// If folder is Exchange and server is offline then skip
			if (pNameSpaceFolder->IsExchange()) {
				Outlook::OlExchangeConnectionMode conn = outlookStorage.GetExchangeConnection();
				if (conn == Outlook::olOffline || conn == Outlook::olDisconnected || conn == Outlook::olCachedDisconnected || conn == Outlook::olCachedOffline) {
					continue;
				}
			}
			
			for (auto pSubFolderInfo : pNameSpaceFolder->GetInboxFolders()) {
				// Make a little pause between processing
				std::this_thread::sleep_for(std::chrono::microseconds(MSG_TIMEOUT));
			
				Outlook::_ItemsPtr pMailItems = nullptr;
				if (pSubFolderInfo.second->get_Items(&pMailItems) != S_OK || pMailItems == nullptr)
					continue;

				IDispatchPtr pDispatch = nullptr;
				auto hr = pMailItems->GetLast(&pDispatch);
				hr = pMailItems->GetFirst(&pDispatch);
				if (pMailItems->GetLast(&pDispatch) != S_OK || pDispatch == nullptr)
					continue;

				DATE newLastMessageTime = pNameSpaceFolder->GetLastMessageTime();

				do {
					Outlook::_MailItemPtr pMailItem(pDispatch);							
					BSTR msgClass = nullptr;
					// Check item class, because it can be SYSTEMMSG, DRAFT or CALENDAR
					if (SUCCEEDED(pMailItem->get_MessageClass(&msgClass)) && msgClass) {
						bool isMail = _wcsicmp(msgClass, IPM_NOTE_MSG_CLASS) == 0;
						SysFreeString(msgClass);
						if (isMail) {
							DATE dt = 0;
							if (SUCCEEDED(pMailItem->get_ReceivedTime(&dt)) && dt > 0)
							{
								if (dt > pNameSpaceFolder->GetLastMessageTime()) {
									// Log message because it's came later then our last message time
									if (COutlookMailProcesser::LogMailItem(pMailItem, true)) {
										// if success then update time
										if (dt > newLastMessageTime) {
											newLastMessageTime = dt;
										}
									}
								}
								// because it's last message, others messages were got earlier
								else {
									break;
								}
							}
						}
					}
				} while (pMailItems->GetPrevious(&pDispatch) == S_OK && pDispatch);
				pNameSpaceFolder->SetLastMessageTime(newLastMessageTime);
			}
		}
		foldersStorage.SaveToFile();

		std::unique_lock<std::mutex> lck(m_mutex);
		m_threadNotifierCv.wait(lck, [this] { return m_hasNewMails || m_workerStopped; });
		m_hasNewMails = false;
	}
	// notify we finished
	m_workerFinishedCv.notify_one();
}

STDMETHODIMP CConnect::OnItemSend(LPDISPATCH Item, VARIANT_BOOL * Cancel)
{
	Outlook::_MailItemPtr pMail = Outlook::_MailItemPtr(Item);
	if (pMail) {
		COutlookMailProcesser::LogMailItem(pMail, false);
	}
	return S_OK;
}

STDMETHODIMP CConnect::OnNewMail()
{
	std::lock_guard<std::mutex> lck(m_mutex);
	m_hasNewMails = true;
	m_threadNotifierCv.notify_all();
	return S_OK;
}
