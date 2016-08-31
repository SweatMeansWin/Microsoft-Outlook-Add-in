#include "PreCompiled.h"
#include "Serializator.h"

bool CSerializator::Serialize(Outlook::_ApplicationPtr pAapplication)
{
	HRESULT hr = ::CreateStreamOnHGlobal(NULL, TRUE, &m_stream);
	if (m_stream) {
		LARGE_INTEGER li = { 0 };
		hr = ::CoMarshalInterface(
			m_stream,
			__uuidof(Outlook::_ApplicationPtr::Interface),
			pAapplication,
			MSHCTX_INPROC,
			NULL,
			MSHLFLAGS_NORMAL
		);
		if (hr == S_OK) {
			hr = m_stream->Seek(li, STREAM_SEEK_SET, NULL);
			return true;
		}
	}
	return false;
}

Outlook::_ApplicationPtr CSerializator::DeSerialize()
{
	Outlook::_ApplicationPtr pointer = nullptr;
	HRESULT hr = ::CoGetInterfaceAndReleaseStream(
		m_stream,
		__uuidof(Outlook::_ApplicationPtr::Interface),
		(void**)&pointer
	);
	return SUCCEEDED(hr) ? pointer : nullptr;
}
