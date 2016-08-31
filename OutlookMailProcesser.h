#pragma once

class COutlookMailProcesser
{
public:
	static bool LogMailItem(Outlook::_MailItemPtr pMail, bool isInputMail);
private:
	static bool saveMailFile(Outlook::_MailItemPtr pMail, const std::wstring& sFileName);
	static bool createMetaFile(Outlook::_MailItemPtr pMail, const std::wstring& sFileName, bool isInputMail);

	static std::wstring m_maildir;
};