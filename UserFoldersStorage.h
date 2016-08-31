#pragma once
#include "PreCompiled.h"
#include "OutlookInstanceStorage.h"
#include "MAPIFolder.h"

class CUserFoldersStorage
{
public:
	CUserFoldersStorage(COutlookInstanceStorage& instanceStorage);
	~CUserFoldersStorage() = default;

	void Add(Outlook::MAPIFolderPtr pMapiFolder);
	bool Contains(Outlook::MAPIFolderPtr pMapiFolder);
	void SaveToFile();

	size_t Size();
	std::shared_ptr<CMAPIFolder> operator[] (size_t index);
private:
	COutlookInstanceStorage& m_instanceStorage;
	std::vector<std::shared_ptr<CMAPIFolder>> m_folders;
	std::wstring m_localListPath;

	void loadFromFile();
	void refreshFolders();
};