#include "PreCompiled.h"

#include "OutlookInstanceStorage.h"

std::wstring DEFAULT_FOLDER_NAME_RU = L"файл данных outlook";
std::wstring DEFAULT_FOLDER_NAME_EN = L"outlook data file";

COutlookInstanceStorage::COutlookInstanceStorage(Outlook::_ApplicationPtr pApplication)
	: m_pApplication(pApplication), m_pNameSpace(nullptr), m_pNSFolders(nullptr), m_bCreated(false)
{
	HRESULT hr = m_pApplication->get_Session(&m_pNameSpace);
	if (SUCCEEDED(hr) && m_pNameSpace) {
		hr = m_pNameSpace->get_Folders(&m_pNSFolders);
		if (SUCCEEDED(hr) && m_pNSFolders) {
			m_bCreated = true;
		}
	}
}

COutlookInstanceStorage::~COutlookInstanceStorage()
{
	if (m_pNSFolders) {
		m_pNSFolders.Release();
		m_pNSFolders = nullptr;
	}
	if (m_pNameSpace) {
		m_pNameSpace.Release();
		m_pNameSpace = nullptr;
	}
	if (m_pApplication) {
		m_pApplication.Release();
		m_pApplication = nullptr;
	}
}

Outlook::_ApplicationPtr COutlookInstanceStorage::GetApplication()
{
	return m_pApplication;
}

Outlook::_NameSpacePtr COutlookInstanceStorage::GetNameSpace()
{
	return m_pNameSpace;
}

Outlook::_FoldersPtr COutlookInstanceStorage::GetNameSpaceFolders()
{
	return m_pNSFolders;
}

Outlook::OlExchangeConnectionMode COutlookInstanceStorage::GetExchangeConnection()
{
	Outlook::OlExchangeConnectionMode conn = Outlook::olNoExchange;
	HRESULT hr = m_pNameSpace->get_ExchangeConnectionMode(&conn);
	return conn;
}

bool COutlookInstanceStorage::IsCreated()
{
	return m_bCreated;
}

bool COutlookInstanceStorage::TryAccess(Outlook::MAPIFolderPtr pMapiFolder) {
	Outlook::_FoldersPtr pFolders = nullptr;
	return SUCCEEDED(pMapiFolder->get_Folders(&pFolders)) && pFolders;
}

bool COutlookInstanceStorage::CheckMailFolder(Outlook::MAPIFolderPtr pFolder)
{
	return COutlookInstanceStorage::TryAccess(pFolder) && COutlookInstanceStorage::IsMailFolder(pFolder) && COutlookInstanceStorage::IsDefaultFolder(pFolder) == false;
}

bool COutlookInstanceStorage::IsDefaultFolder(Outlook::MAPIFolderPtr pAccountFolder) {
	std::wstring folderName = GetFolderName(pAccountFolder);
	return folderName == DEFAULT_FOLDER_NAME_RU || folderName == DEFAULT_FOLDER_NAME_EN;
}

std::wstring COutlookInstanceStorage::GetFolderName(Outlook::MAPIFolderPtr pMapiFolder) {
	std::wstring sFolderName;
	BSTR szFolderName = nullptr;
	if (SUCCEEDED(pMapiFolder->get_Name(&szFolderName)) && szFolderName) {
		sFolderName.append(szFolderName);
		SysFreeString(szFolderName);

		std::transform(sFolderName.begin(), sFolderName.end(), sFolderName.begin(), towlower);
	}
	return sFolderName;
}

std::wstring COutlookInstanceStorage::GetEntryID(Outlook::MAPIFolderPtr pMapiFolder) {
	std::wstring sEntryId;
	BSTR szEntryID = nullptr;
	if (SUCCEEDED(pMapiFolder->get_EntryID(&szEntryID)) && szEntryID) {
		sEntryId.append(szEntryID);
		SysFreeString(szEntryID);
	}
	return sEntryId;
}

bool COutlookInstanceStorage::IsMailFolder(Outlook::MAPIFolderPtr pFolder) {
	Outlook::OlItemType itemType = Outlook::olMailItem;
	return SUCCEEDED(pFolder->get_DefaultItemType(&itemType)) && itemType == Outlook::olMailItem;
}