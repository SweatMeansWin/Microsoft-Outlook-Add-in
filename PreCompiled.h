// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#define NOMINMAX
#import "MSADDNDR.OLB" raw_interfaces_only, raw_native_types, no_namespace, named_guids, auto_search
#import "MSOUTL.OLB" raw_interfaces_only, raw_native_types, named_guids, auto_search, \
	rename("CopyFile","OutlCopyFile"), rename("PlaySound","OutlPlaySound"), rename("Folder","OutlFolder"), \
	rename("DocumentProperties","OutlDocumentProperties"), rename("RGB","OutlRGB")

// Declarations for different library platforms
#ifdef x64
#define LIBID_OutlookAddinLib LIBID_OutlookAddinLib64
#define INTERFACE_HEADER "OutlookAddinProject64_i.h"
#define CLSID_Connect CLSID_Connect64
#define IID_IConnect IID_IConnect64
#define IConnect IConnect64
#define ConnectClass Connect64
#define REGISTRY_GUID "{7185F9A3-BBEC-48A0-ACC7-44FA5ED8CDD8}"
#else
#define LIBID_OutlookAddinLib LIBID_OutlookAddinLib32
#define INTERFACE_HEADER "OutlookAddinProject32_i.h"
#define CLSID_Connect CLSID_Connect32
#define IID_IConnect IID_IConnect32
#define IConnect IConnect32
#define ConnectClass Connect32
#define REGISTRY_GUID "{A0303C54-34E0-4136-BDBF-2BCEE4684C79}"
#endif

#include <SDKDDKVer.h>

#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE
#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit
#define ATL_NO_ASSERT_ON_DESTROY_NONEXISTENT_WINDOW

#include "resource.h"

#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>

#include <algorithm>
#include <atomic>
#include <fstream>
#include <functional>
#include <future>
#include <memory>
#include <mutex>
#include <set>
#include <sstream>
#include <vector>
