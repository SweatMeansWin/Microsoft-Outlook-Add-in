#include "PreCompiled.h"
#include "UserFoldersStorage.h"
#include "OutlookInstanceStorage.h"

wchar_t DELIMITER = L':';

CUserFoldersStorage::CUserFoldersStorage(COutlookInstanceStorage& instance)
	: m_instanceStorage(instance), m_localListPath(L"tmp")
{
	refreshFolders();
	loadFromFile();
}

void CUserFoldersStorage::Add(Outlook::MAPIFolderPtr pMapiFolder)
{
	m_folders.push_back(std::make_shared<CMAPIFolder>(pMapiFolder));
}

bool CUserFoldersStorage::Contains(Outlook::MAPIFolderPtr pMapiFolder)
{
	auto cit = std::find_if(m_folders.cbegin(), m_folders.cend(), [&pMapiFolder](std::shared_ptr<CMAPIFolder> cit) {
		return cit->CheckEqual(pMapiFolder);
	});
	return cit != m_folders.cend();
}

void CUserFoldersStorage::refreshFolders()
{
	auto pFolders = m_instanceStorage.GetNameSpaceFolders();
	if (pFolders) {
		Outlook::MAPIFolderPtr pMapiFolder = nullptr;
		if (SUCCEEDED(pFolders->GetFirst(&pMapiFolder)) && pMapiFolder) {
			do {
				if (COutlookInstanceStorage::IsMailFolder(pMapiFolder)) {
					Add(pMapiFolder);
				}
			} while (SUCCEEDED(pFolders->GetNext(&pMapiFolder)) && pMapiFolder);
		}
	}
}

void CUserFoldersStorage::loadFromFile()
{
	std::wifstream wif(m_localListPath);
	if (!wif.is_open()) {
		return;
	}
	std::wstring sFolderTimeLine;
	while (std::getline(wif, sFolderTimeLine)) {
		auto delim_pos = sFolderTimeLine.find(DELIMITER);

		for (size_t idx = 0; idx < m_folders.size(); ++idx) {
			auto pointer = m_folders.at(idx);
			if (pointer->GetFolderName() == sFolderTimeLine.substr(0, delim_pos)) {
				pointer->SetLastMessageTime(std::stod(sFolderTimeLine.substr(delim_pos + 1)));
			}
		}
	}
}

void CUserFoldersStorage::SaveToFile()
{
	std::wofstream wof(m_localListPath);
	for (size_t idx = 0; idx < m_folders.size(); ++idx) {
		auto pointer = m_folders.at(idx);
		wof << pointer->GetFolderName() << DELIMITER << pointer->GetLastMessageTime() << std::endl;
	}
}

size_t CUserFoldersStorage::Size()
{
	return m_folders.size();
}

std::shared_ptr<CMAPIFolder> CUserFoldersStorage::operator[](size_t index)
{
	if (index < m_folders.size()) {
		return m_folders.at(index);
	}
	return nullptr;
}
