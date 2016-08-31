#pragma once


class CSerializator
{
public:
	CSerializator() : m_stream(nullptr) {}
	~CSerializator() = default;

	bool Serialize(Outlook::_ApplicationPtr pAapplication);
	Outlook::_ApplicationPtr DeSerialize();

private:
	IStream * m_stream;
};

