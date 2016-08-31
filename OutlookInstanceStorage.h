#pragma once

class COutlookInstanceStorage
{
public:
	COutlookInstanceStorage(Outlook::_ApplicationPtr pApplication);
	~COutlookInstanceStorage();

	Outlook::_ApplicationPtr GetApplication();
	Outlook::_NameSpacePtr GetNameSpace();
	Outlook::_FoldersPtr GetNameSpaceFolders();
	Outlook::OlExchangeConnectionMode GetExchangeConnection();
	bool IsCreated();

	static bool CheckMailFolder(Outlook::MAPIFolderPtr pFolder);
	static bool	IsDefaultFolder(Outlook::MAPIFolderPtr pAccountFolder);
	static bool	TryAccess(Outlook::MAPIFolderPtr pMapiFolder);
	static bool	IsMailFolder(Outlook::MAPIFolderPtr pMapiFolder);
	static std::wstring GetFolderName(Outlook::MAPIFolderPtr pMapiFolder);
	static std::wstring GetEntryID(Outlook::MAPIFolderPtr pMapiFolder);
private:
	Outlook::_ApplicationPtr m_pApplication;
	Outlook::_NameSpacePtr m_pNameSpace;
	Outlook::_FoldersPtr m_pNSFolders;
	bool m_bCreated;
};