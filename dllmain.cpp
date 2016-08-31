// dllmain.cpp : Implementation of DllMain.

#include "PreCompiled.h"
#include INTERFACE_HEADER
#include "dllmain.h"

COutlookAddinProjectModule _AtlModule;

// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	hInstance;
	return _AtlModule.DllMain(dwReason, lpReserved);
}
