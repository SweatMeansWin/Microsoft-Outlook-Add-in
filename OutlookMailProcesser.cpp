#include "PreCompiled.h"
#include "OutlookMailProcesser.h"

#include "Utf8Codecvt.h"
#include "Utils.h"
#include <Sddl.h>

const int DOWNLOAD_MAIL_RETRY_COUNT = 10;
const std::chrono::seconds DOWNLOAD_MAIL_TIMEOUT_SEC(1);

std::wstring COutlookMailProcesser::m_maildir;

BSTR BODY_PROPERTY = L"http://schemas.microsoft.com/mapi/proptag/0x1000001E";
BSTR SUBJECT_PROPERTY = L"http://schemas.microsoft.com/mapi/proptag/0x0037001E";
BSTR SENDER_EMAIL_PROPERTY = L"http://schemas.microsoft.com/mapi/proptag/0x0C1F001E";

bool COutlookMailProcesser::LogMailItem(Outlook::_MailItemPtr pMailItem, bool isInputMail)
{
	// Download input mail local in outlook to save it to FS
	if (isInputMail) {
		Outlook::OlDownloadState downloadState = Outlook::olHeaderOnly;
		if (SUCCEEDED(pMailItem->get_DownloadState(&downloadState)) && downloadState != Outlook::olFullItem) {
			// Try to download
			if (SUCCEEDED(pMailItem->put_MarkForDownload(Outlook::olMarkedForDownload))) {
				for (int idx = 0; idx < DOWNLOAD_MAIL_RETRY_COUNT; ++idx) {
					if (SUCCEEDED(pMailItem->get_DownloadState(&downloadState)) && downloadState != Outlook::olFullItem) {
						std::this_thread::sleep_for(DOWNLOAD_MAIL_TIMEOUT_SEC);
					}
				}
			}
			pMailItem->Save();
		}
	}

	SYSTEMTIME localTime = { 0 };
	GetLocalTime(&localTime);

	std::wstringstream wss;
	wss << m_maildir << L"\\" << CTimeOperations::GetTimeStamp(localTime);
	std::wstring sMsgFileName = wss.str() + L".msg";
	std::wstring sMetaFileName = wss.str() + L".meta";

	bool success = false;
	if (success = saveMailFile(pMailItem, sMsgFileName)) {
		createMetaFile(pMailItem, sMetaFileName, isInputMail);
	}
	if (!success) {
		DeleteFile(sMsgFileName.c_str());
		DeleteFile(sMetaFileName.c_str());
	}
	return success;
}

bool COutlookMailProcesser::saveMailFile(Outlook::_MailItemPtr pMailItem, const std::wstring& sFileName)
{
	BSTR str = SysAllocStringLen(sFileName.c_str(), static_cast<UINT>(sFileName.size() * sizeof(wchar_t)));
	if (str) {
		HRESULT hr = pMailItem->SaveAs(str, _variant_t(Outlook::olMSG));
		SysFreeString(str);
		return SUCCEEDED(hr);
	}
	return false;
}

bool COutlookMailProcesser::createMetaFile(Outlook::_MailItemPtr pMail, const std::wstring& FileName, bool isInputMail)
{
	bool success = false;
	const int SID_MAX_SIZE = 1024;

	// User
	wchar_t szUserName[MAX_PATH] = { 0 };
	DWORD dwUserCb = MAX_PATH;
	GetUserName(szUserName, &dwUserCb);
	// Domain
	wchar_t szDomainName[MAX_PATH] = { 0 };
	DWORD dwDomainCb = MAX_PATH;

	SID_NAME_USE sidNameUse;
	BYTE pBytes[SID_MAX_SIZE] = { 0 };
	PSID pSid = (PSID)pBytes;
	DWORD dwSidCb = SID_MAX_SIZE;
	if (LookupAccountName(NULL, szUserName, pSid, &dwSidCb, szDomainName, &dwDomainCb, &sidNameUse)) {
		wchar_t sSid[SID_MAX_SIZE] = { 0 };
		wchar_t* beg = &sSid[0];
		ConvertSidToStringSid(pSid, &beg);
		std::wofstream metaFile;
		metaFile.imbue(stdx::get_utf8_locale());
		metaFile.open(FileName, std::ios_base::out | std::ios_base::binary);
		if (!metaFile.is_open()) {
			return false;
		}

		metaFile << L"OUT:" << (isInputMail ? L"0" : L"1") << std::endl;
		metaFile << L"SID:" << sSid << std::endl;
		metaFile << L"User:" << szDomainName << L"\\" << szUserName << std::endl;

		Outlook::_PropertyAccessorPtr pPropAcc;
		pMail->get_PropertyAccessor(&pPropAcc);

		VARIANT value;
		VariantInit(&value);

		metaFile << L"Subject:";
		{
			BSTR sSubject = nullptr;
			if (SUCCEEDED(pMail->get_Subject(&sSubject)) && sSubject) {
				metaFile << sSubject;
				SysFreeString(sSubject);
			}
			else {
				if (SUCCEEDED(pPropAcc->GetProperty(SUBJECT_PROPERTY, &value))) {
					metaFile << value.bstrVal;
				}
			}
			VariantClear(&value);
		}
		metaFile << std::endl;

		metaFile << L"From:";
		{
			BSTR szSenderEmail = nullptr;
			bool valueIsSet = false;

			if (!isInputMail) {
				Outlook::_AccountPtr pAccount = nullptr;
				if (SUCCEEDED(pMail->get_SendUsingAccount(&pAccount)) && pAccount) {
					if (SUCCEEDED(pAccount->get_SmtpAddress(&szSenderEmail)) && szSenderEmail) {
						metaFile << szSenderEmail;
						SysFreeString(szSenderEmail);
						valueIsSet = true;
					}
				}
				else {
					if (SUCCEEDED(pPropAcc->GetProperty(SENDER_EMAIL_PROPERTY, &value))) {
						metaFile << value.bstrVal;
					}
					VariantClear(&value);
				}
			}

			if (!valueIsSet)
			{
				Outlook::AddressEntryPtr pAddEntry = nullptr;
				if (SUCCEEDED(pMail->get_Sender(&pAddEntry)) && pAddEntry) {
					metaFile << GetAccountAddress(pAddEntry);
					valueIsSet = true;
				}
				else {
					if (SUCCEEDED(pMail->get_SenderEmailAddress(&szSenderEmail)) && szSenderEmail) {
						metaFile << szSenderEmail;
						SysFreeString(szSenderEmail);
						valueIsSet = true;
					}
				}
			}
		}
		metaFile << std::endl;

		metaFile << L"Sent:";
		{
			DATE dtSentDate = 0;
			if (SUCCEEDED(pMail->get_SentOn(&dtSentDate)) && dtSentDate > 0) {
				SYSTEMTIME st = { 0 };
				GetLocalTime(&st);

				DATE dtNow = CTimeOperations::ConvertSystemTimeToDate(st);
				if (dtSentDate >= dtNow)
					metaFile << CTimeOperations::GetTimeStamp(st);
				else
					metaFile << CTimeOperations::GetTimeStamp(CTimeOperations::ConvertDateToSystemTime(dtSentDate));
			}
		}
		metaFile << std::endl;

		metaFile << L"To:";
		{
			Outlook::RecipientsPtr pRecipients = nullptr;
			if (SUCCEEDED(pMail->get_Recipients(&pRecipients)) && pRecipients) {
				long lCount = 0;
				pRecipients->get_Count(&lCount);
				if (lCount > 0) {
					for (long i = 1; i <= lCount; ++i) {
						Outlook::RecipientPtr pRecipient = nullptr;
						if (SUCCEEDED(pRecipients->Item(_variant_t(i), &pRecipient)) && pRecipient) {
							Outlook::AddressEntryPtr pAddEntry = nullptr;
							if (SUCCEEDED(pRecipient->get_AddressEntry(&pAddEntry)) && pAddEntry) {
								metaFile << GetAccountAddress(pAddEntry);
								if (i <= lCount - 1)
									metaFile << L" ";
							}
						}
					}
				}
			}
		}
		metaFile << std::endl;

		metaFile << L"CC:";
		{
			BSTR szCC = nullptr;
			if (SUCCEEDED(pMail->get_CC(&szCC)) && szCC) {
				metaFile << szCC;
				SysFreeString(szCC);
			}
		}
		metaFile << std::endl;

		metaFile << L"BC:";
		{
			BSTR szBCC = nullptr;
			if (SUCCEEDED(pMail->get_BCC(&szBCC)) && szBCC) {
				metaFile << szBCC;
				SysFreeString(szBCC);
			}
		}
		metaFile << std::endl;

		//Attachments
		{
			Outlook::AttachmentsPtr pAttachments = nullptr;
			if (SUCCEEDED(pMail->get_Attachments(&pAttachments)) && pAttachments) {
				metaFile << L"Attcount:";
				long lCount = 0;
				pAttachments->get_Count(&lCount);
				metaFile << lCount;
				metaFile << std::endl;

				//Enum them
				for (long i = 1; i <= lCount; i++) {
					metaFile << i << L":";

					Outlook::AttachmentPtr pAttachment = nullptr;
					if (SUCCEEDED(pAttachments->Item(_variant_t(i), &pAttachment)) && pAttachment) {
						BSTR szAttachName = nullptr;
						if (SUCCEEDED(pAttachment->get_FileName(&szAttachName)) && szAttachName) {
							metaFile << szAttachName;
							SysFreeString(szAttachName);
						}
					}
					metaFile << std::endl;
				}
			}
		}

		metaFile << L"Body:";
		{
			BSTR szBody = nullptr;
			if (SUCCEEDED(pMail->get_Body(&szBody)) && szBody) {
				metaFile << szBody;
				SysFreeString(szBody);
			}
			else
			{
				if (SUCCEEDED(pPropAcc->GetProperty(BODY_PROPERTY, &value)) && value.bstrVal) {
					metaFile << value.bstrVal;
				}
				VariantClear(&value);
			}
		}
		metaFile << std::endl;
		success = true;
	}

	return success;
}
