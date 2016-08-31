

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0622 */
/* at Tue Jan 19 08:14:07 2038
 */
/* Compiler settings for OutlookAddinProject32.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.01.0622 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */



/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 500
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif /* __RPCNDR_H_VERSION__ */

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __OutlookAddinProject32_i_h__
#define __OutlookAddinProject32_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IConnect32_FWD_DEFINED__
#define __IConnect32_FWD_DEFINED__
typedef interface IConnect32 IConnect32;

#endif 	/* __IConnect32_FWD_DEFINED__ */


#ifndef __Connect32_FWD_DEFINED__
#define __Connect32_FWD_DEFINED__

#ifdef __cplusplus
typedef class Connect32 Connect32;
#else
typedef struct Connect32 Connect32;
#endif /* __cplusplus */

#endif 	/* __Connect32_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IConnect32_INTERFACE_DEFINED__
#define __IConnect32_INTERFACE_DEFINED__

/* interface IConnect32 */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IConnect32;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("ECD05AE6-3412-4407-A6BA-1298AB865916")
    IConnect32 : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IConnect32Vtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IConnect32 * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IConnect32 * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IConnect32 * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IConnect32 * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IConnect32 * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IConnect32 * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IConnect32 * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } IConnect32Vtbl;

    interface IConnect32
    {
        CONST_VTBL struct IConnect32Vtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IConnect32_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IConnect32_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IConnect32_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IConnect32_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IConnect32_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IConnect32_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IConnect32_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IConnect32_INTERFACE_DEFINED__ */



#ifndef __OutlookAddinLib32_LIBRARY_DEFINED__
#define __OutlookAddinLib32_LIBRARY_DEFINED__

/* library OutlookAddinLib32 */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_OutlookAddinLib32;

EXTERN_C const CLSID CLSID_Connect32;

#ifdef __cplusplus

class DECLSPEC_UUID("2B274807-B9AC-471A-B643-68C0FDFD5CB0")
Connect32;
#endif
#endif /* __OutlookAddinLib32_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


