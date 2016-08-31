#include "PreCompiled.h"
#include "Utils.h"

DATE CTimeOperations::ConvertSystemTimeToDate(SYSTEMTIME& st)
{
	DATE dt = 0;
	SystemTimeToVariantTime(&st, &dt);
	return dt;
}

SYSTEMTIME CTimeOperations::ConvertDateToSystemTime(DATE dt)
{
	SYSTEMTIME st = { 0 };
	VariantTimeToSystemTime(dt, &st);
	return st;
}

std::wstring CTimeOperations::GetTimeStamp(SYSTEMTIME & st)
{
	wchar_t time[MAX_PATH] = { 0 };
	_stprintf_s(time, MAX_PATH,
		_T("%d-%02d-%02d-%02d-%02d-%02d-%03d"),
		st.wYear,
		st.wMonth,
		st.wDay,
		st.wHour,
		st.wMinute,
		st.wSecond,
		st.wMilliseconds);
	return time;
}

std::wstring GetAccountAddress(Outlook::AddressEntryPtr pAddressEntry)
{
	std::wstring sAccountAddress;
	BSTR szSmtpAddress = nullptr;

	Outlook::OlAddressEntryUserType userType;
	if (SUCCEEDED(pAddressEntry->get_AddressEntryUserType(&userType)) && userType == Outlook::olSmtpAddressEntry) {
		if (SUCCEEDED(pAddressEntry->get_Address(&szSmtpAddress)) && szSmtpAddress) {
		}
	}
	else {
		Outlook::_ExchangeUserPtr pExchUser = nullptr;
		if (pAddressEntry->GetExchangeUser(&pExchUser) == S_OK && pExchUser) {
			if (SUCCEEDED(pExchUser->get_PrimarySmtpAddress(&szSmtpAddress)) && szSmtpAddress) {
				sAccountAddress.append(szSmtpAddress);
			}
			else if (SUCCEEDED(pExchUser->get_Address(&szSmtpAddress)) && szSmtpAddress) {
				sAccountAddress.append(szSmtpAddress);
			}
		}
		else if (SUCCEEDED(pAddressEntry->get_Address(&szSmtpAddress)) && szSmtpAddress) {
			sAccountAddress.append(szSmtpAddress);
		}
	}
	if (szSmtpAddress != nullptr) {
		SysFreeString(szSmtpAddress);
	}
	return sAccountAddress;
}
