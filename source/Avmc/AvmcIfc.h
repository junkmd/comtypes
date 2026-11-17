

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0628 */
/* at Tue Jan 19 03:14:07 2038
 */
/* Compiler settings for AvmcIfc.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=AMD64 8.01.0628 
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

#ifndef __AvmcIfc_h__
#define __AvmcIfc_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef DECLSPEC_XFGVIRT
#if defined(_CONTROL_FLOW_GUARD_XFG)
#define DECLSPEC_XFGVIRT(base, func) __declspec(xfg_virtual(base, func))
#else
#define DECLSPEC_XFGVIRT(base, func)
#endif
#endif

/* Forward Declarations */ 

#ifndef __IAvmc_FWD_DEFINED__
#define __IAvmc_FWD_DEFINED__
typedef interface IAvmc IAvmc;

#endif 	/* __IAvmc_FWD_DEFINED__ */


#ifndef __Avmc_FWD_DEFINED__
#define __Avmc_FWD_DEFINED__

#ifdef __cplusplus
typedef class Avmc Avmc;
#else
typedef struct Avmc Avmc;
#endif /* __cplusplus */

#endif 	/* __Avmc_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_AvmcIfc_0000_0000 */
/* [local] */ 

typedef /* [helpstring][version][uuid] */  DECLSPEC_UUID("6C7A25CB-7938-4BE0-A285-12C616717FDD") struct DeviceInfo
    {
    /* [helpstring] */ VARIANT Special;
    /* [helpstring] */ BSTR Name;
    /* [helpstring] */ long Value;
    /* [helpstring] */ long Flags;
    /* [helpstring] */ long Type;
    /* [helpstring] */ long ID;
    /* [helpstring] */ long LocId;
    /* [helpstring] */ BSTR SerialNumber;
    /* [helpstring] */ BSTR Description;
    /* [helpstring] */ long ftHandle;
    } 	DeviceInfo;



extern RPC_IF_HANDLE __MIDL_itf_AvmcIfc_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_AvmcIfc_0000_0000_v0_0_s_ifspec;

#ifndef __IAvmc_INTERFACE_DEFINED__
#define __IAvmc_INTERFACE_DEFINED__

/* interface IAvmc */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IAvmc;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("6C7A25CC-7938-4BE0-A285-12C616717FDD")
    IAvmc : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE FindAllAvmc( 
            /* [out] */ SAFEARRAY * *avmcList) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IAvmcVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IAvmc * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IAvmc * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IAvmc * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IAvmc * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IAvmc * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IAvmc * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IAvmc * This,
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
        
        DECLSPEC_XFGVIRT(IAvmc, FindAllAvmc)
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *FindAllAvmc )( 
            IAvmc * This,
            /* [out] */ SAFEARRAY * *avmcList);
        
        END_INTERFACE
    } IAvmcVtbl;

    interface IAvmc
    {
        CONST_VTBL struct IAvmcVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IAvmc_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IAvmc_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IAvmc_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IAvmc_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IAvmc_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IAvmc_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IAvmc_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IAvmc_FindAllAvmc(This,avmcList)	\
    ( (This)->lpVtbl -> FindAllAvmc(This,avmcList) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IAvmc_INTERFACE_DEFINED__ */



#ifndef __AVMCIFCLib_LIBRARY_DEFINED__
#define __AVMCIFCLib_LIBRARY_DEFINED__

/* library AVMCIFCLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_AVMCIFCLib;

EXTERN_C const CLSID CLSID_Avmc;

#ifdef __cplusplus

class DECLSPEC_UUID("41BDBDFC-A848-4523-A149-ADD3AE1E6D84")
Avmc;
#endif
#endif /* __AVMCIFCLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  LPSAFEARRAY_UserSize(     unsigned long *, unsigned long            , LPSAFEARRAY * ); 
unsigned char * __RPC_USER  LPSAFEARRAY_UserMarshal(  unsigned long *, unsigned char *, LPSAFEARRAY * ); 
unsigned char * __RPC_USER  LPSAFEARRAY_UserUnmarshal(unsigned long *, unsigned char *, LPSAFEARRAY * ); 
void                      __RPC_USER  LPSAFEARRAY_UserFree(     unsigned long *, LPSAFEARRAY * ); 

unsigned long             __RPC_USER  LPSAFEARRAY_UserSize64(     unsigned long *, unsigned long            , LPSAFEARRAY * ); 
unsigned char * __RPC_USER  LPSAFEARRAY_UserMarshal64(  unsigned long *, unsigned char *, LPSAFEARRAY * ); 
unsigned char * __RPC_USER  LPSAFEARRAY_UserUnmarshal64(unsigned long *, unsigned char *, LPSAFEARRAY * ); 
void                      __RPC_USER  LPSAFEARRAY_UserFree64(     unsigned long *, LPSAFEARRAY * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


