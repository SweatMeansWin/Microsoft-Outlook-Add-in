#pragma once

class CTimeOperations
{
public:
	static DATE	CTimeOperations::ConvertSystemTimeToDate(SYSTEMTIME& st);
	static SYSTEMTIME CTimeOperations::ConvertDateToSystemTime(DATE dt);
	static std::wstring	CTimeOperations::GetTimeStamp(SYSTEMTIME& st);
};

std::wstring GetAccountAddress(Outlook::AddressEntryPtr pAddressEntry);
