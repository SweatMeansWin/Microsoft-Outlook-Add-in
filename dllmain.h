// dllmain.h : Declaration of module class.

class COutlookAddinProjectModule : public ATL::CAtlDllModuleT< COutlookAddinProjectModule >
{
public :
	DECLARE_LIBID(LIBID_OutlookAddinLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_OUTLOOKADDINPROJECT, REGISTRY_GUID)
};

extern class COutlookAddinProjectModule _AtlModule;
