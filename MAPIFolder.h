#pragma once

class CMAPIFolder
{
public:
	CMAPIFolder(Outlook::MAPIFolderPtr pMapiFolder, DATE dtLastMessageTime=0);
	~CMAPIFolder() = default;

	bool IsExchange();
	bool CheckEqual(Outlook::MAPIFolderPtr pMapiFolder);

	void SetLastMessageTime(DATE time);
	DATE GetLastMessageTime();
	const std::wstring& GetFolderName();
	Outlook::MAPIFolderPtr GetCOMInstance();
	std::vector<std::pair<std::wstring, Outlook::MAPIFolderPtr>> GetInboxFolders();
private:
	Outlook::MAPIFolderPtr m_pMapiFolder;
	std::vector<std::pair<std::wstring, Outlook::MAPIFolderPtr>> m_inboxFoldersVec;
	std::set<std::wstring> m_ignoringFolders;
	std::wstring m_sFolderName;
	DATE m_dtLastMessageTime;

	void setLastMessageTime();
	void setFolders();
};
