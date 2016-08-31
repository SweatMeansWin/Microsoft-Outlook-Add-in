

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0622 */
/* at Tue Jan 19 08:14:07 2038
 */
/* Compiler settings for OutlookAddinProject64.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=AMD64 8.01.0622 
    protocol : all , ms_ext, c_ext, robust
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

#ifndef __OutlookAddinProject64_i_h__
#define __OutlookAddinProject64_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IConnect64_FWD_DEFINED__
#define __IConnect64_FWD_DEFINED__
typedef interface IConnect64 IConnect64;

#endif 	/* __IConnect64_FWD_DEFINED__ */


#ifndef __Connect64_FWD_DEFINED__
#define __Connect64_FWD_DEFINED__

#ifdef __cplusplus
typedef class Connect64 Connect64;
#else
typedef struct Connect64 Connect64;
#endif /* __cplusplus */

#endif 	/* __Connect64_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IConnect64_INTERFACE_DEFINED__
#define __IConnect64_INTERFACE_DEFINED__

/* interface IConnect64 */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IConnect64;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("5933C7D6-C625-4B06-8009-A9147684CD22")
    IConnect64 : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IConnect64Vtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IConnect64 * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IConnect64 * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IConnect64 * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IConnect64 * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IConnect64 * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IConnect64 * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IConnect64 * This,
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
    } IConnect64Vtbl;

    interface IConnect64
    {
        CONST_VTBL struct IConnect64Vtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IConnect64_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IConnect64_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IConnect64_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IConnect64_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IConnect64_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IConnect64_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IConnect64_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IConnect64_INTERFACE_DEFINED__ */



#ifndef __OutlookAddinLib64_LIBRARY_DEFINED__
#define __OutlookAddinLib64_LIBRARY_DEFINED__

/* library OutlookAddinLib64 */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_OutlookAddinLib64;

EXTERN_C const CLSID CLSID_Connect64;

#ifdef __cplusplus

class DECLSPEC_UUID("3C07CFA4-99FC-423B-8983-87AE8F40D3DF")
Connect64;
#endif
#endif /* __OutlookAddinLib64_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


