#include "PreCompiled.h"
#include "MAPIFolder.h"
#include "OutlookInstanceStorage.h"
#include "Utils.h"

const int defaultFolders[] = {
	3,//olFolderDeletedItems
	4,//olFolderOutbox
	5,//olFolderSentMail
	9,//olFolderCalendar
	10,//olFolderContacts
	11,//olFolderJournal
	12,//olFolderNotes
	13,//olFolderTasks
	16,//olFolderDrafts
	18,//olPublicFoldersAllPublicFolders
	19,//olFolderConflicts
	20,//olFolderSyncIssues
	21,//olFolderLocalFailures
	22,//olFolderServerFailures
	23,//olFolderJunk
	25,//olFolderRssFeeds
	28,//olFolderToDo
	29,//olFolderManagedEmail
	30//olFolderSuggestedContacts
	  //except 6,//olFolderInbox 
};
LPWSTR FOLDER_SUFFIX = L"this computer only";


CMAPIFolder::CMAPIFolder(Outlook::MAPIFolderPtr pMapiFolder, DATE dtLastMessageTime)
	: m_pMapiFolder(pMapiFolder)
{
	if (dtLastMessageTime > 0) {
		m_dtLastMessageTime = dtLastMessageTime;
	}
	else {
		m_dtLastMessageTime = 0;
		setLastMessageTime();
	}
	m_sFolderName = COutlookInstanceStorage::GetFolderName(pMapiFolder);	
	setFolders();
}

bool CMAPIFolder::IsExchange()
{
	bool IsExchange = false;

	Outlook::_StorePtr pStore = nullptr;
	HRESULT hr = m_pMapiFolder->get_Store(&pStore);
	if (SUCCEEDED(hr) && pStore) {
		Outlook::OlExchangeStoreType exStoreType = Outlook::olNotExchange;
		if (SUCCEEDED(pStore->get_ExchangeStoreType(&exStoreType))) {
			IsExchange = exStoreType != Outlook::olNotExchange;
		}
	}
	return IsExchange;
}

bool CMAPIFolder::CheckEqual(Outlook::MAPIFolderPtr pMapiFolder)
{
	return COutlookInstanceStorage::GetFolderName(pMapiFolder) == m_sFolderName;
}

void CMAPIFolder::SetLastMessageTime(DATE time)
{
	m_dtLastMessageTime = time;
}

DATE CMAPIFolder::GetLastMessageTime()
{
	return m_dtLastMessageTime;
}

const std::wstring & CMAPIFolder::GetFolderName()
{
	return m_sFolderName;
}

Outlook::MAPIFolderPtr CMAPIFolder::GetCOMInstance()
{
	return m_pMapiFolder;
}

std::vector<std::pair<std::wstring, Outlook::MAPIFolderPtr>> CMAPIFolder::GetInboxFolders()
{
	return m_inboxFoldersVec;
}

void CMAPIFolder::setLastMessageTime()
{
	Outlook::_FoldersPtr pSubFolders = nullptr;
	// Get all subfolders collection
	if (SUCCEEDED(m_pMapiFolder->get_Folders(&pSubFolders)) && pSubFolders) {
		Outlook::MAPIFolderPtr pSubFolder = nullptr;
		// Get first subfolder and then enum
		if (SUCCEEDED(pSubFolders->GetFirst(&pSubFolder)) && pSubFolder) {
			do {
				// Check folders default itemtype, it could be calendar or system logs
				if (COutlookInstanceStorage::IsMailFolder(pSubFolder)) {
					//Walk through all folders inbox/trash/etc
					Outlook::_ItemsPtr pItems = nullptr;
					if (SUCCEEDED(pSubFolder->get_Items(&pItems)) && pItems) {
						long lCount = 0;
						if (SUCCEEDED(pItems->get_Count(&lCount)) && lCount > 0) {
							IDispatchPtr pDispatch = nullptr;
							DATE tempReceivedTime = 0;

							// Check first and last elements time, because different folders can have reversed sort order
							if (SUCCEEDED(pItems->GetFirst(&pDispatch)) && pDispatch) {
								if (SUCCEEDED(Outlook::_MailItemPtr(pDispatch)->get_ReceivedTime(&tempReceivedTime))) {
									if (tempReceivedTime > m_dtLastMessageTime) {
										m_dtLastMessageTime = tempReceivedTime;
									}
								}
							}
							if (SUCCEEDED(pItems->GetLast(&pDispatch)) && pDispatch) {
								if (SUCCEEDED(Outlook::_MailItemPtr(pDispatch)->get_ReceivedTime(&tempReceivedTime))) {
									if (tempReceivedTime > m_dtLastMessageTime) {
										m_dtLastMessageTime = tempReceivedTime;
									}
								}
							};
						}
					}
				}
			} while (SUCCEEDED(pSubFolders->GetNext(&pSubFolder)) && pSubFolder);
		}
	}
	// If we have problems or don't have items then store current time
	if (m_dtLastMessageTime == 0) {
		SYSTEMTIME st = { 0 };
		GetLocalTime(&st);
		m_dtLastMessageTime = CTimeOperations::ConvertSystemTimeToDate(st);
	}
}

void CMAPIFolder::setFolders()
{
	Outlook::_StorePtr pStore = nullptr;
	if (SUCCEEDED(m_pMapiFolder->get_Store(&pStore)) && pStore) {
		// Iterate all defaults folders and put them to ignore
		for each (int defaultFolder in defaultFolders) {
			Outlook::MAPIFolderPtr pNotInboxFolder = nullptr;
			HRESULT hr = pStore->GetDefaultFolder(static_cast<Outlook::OlDefaultFolders>(defaultFolder), &pNotInboxFolder);
			if (SUCCEEDED(hr) && pNotInboxFolder) {
				m_ignoringFolders.insert(COutlookInstanceStorage::GetFolderName(pNotInboxFolder));
			}
		}
	}

	Outlook::_FoldersPtr pSubFolders = nullptr;
	// Enumarate all MAPI folders and add non-ignored
	if (SUCCEEDED(m_pMapiFolder->get_Folders(&pSubFolders)) && pSubFolders) {
		Outlook::MAPIFolderPtr pSubFolder = nullptr;
		if (SUCCEEDED(pSubFolders->GetFirst(&pSubFolder)) && pSubFolder) {
			do {
				if (COutlookInstanceStorage::IsMailFolder(pSubFolder)) {
					std::wstring sFolderName = COutlookInstanceStorage::GetFolderName(pSubFolder);
					if (m_ignoringFolders.find(sFolderName) == m_ignoringFolders.end() && sFolderName.find(FOLDER_SUFFIX) == sFolderName.npos) {
						m_inboxFoldersVec.push_back(std::make_pair(sFolderName, pSubFolder));
					}
				}
			} while (SUCCEEDED(pSubFolders->GetNext(&pSubFolder)) && pSubFolder);
		}
	}
}
