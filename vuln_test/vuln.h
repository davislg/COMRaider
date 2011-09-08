/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Thu Feb 23 18:26:14 2006
 */
/* Compiler settings for D:\work_data\COMRaider\vuln_test\vuln.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __vuln_h__
#define __vuln_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __Iserver_FWD_DEFINED__
#define __Iserver_FWD_DEFINED__
typedef interface Iserver Iserver;
#endif 	/* __Iserver_FWD_DEFINED__ */


#ifndef ___IserverEvents_FWD_DEFINED__
#define ___IserverEvents_FWD_DEFINED__
typedef interface _IserverEvents _IserverEvents;
#endif 	/* ___IserverEvents_FWD_DEFINED__ */


#ifndef __server_FWD_DEFINED__
#define __server_FWD_DEFINED__

#ifdef __cplusplus
typedef class server server;
#else
typedef struct server server;
#endif /* __cplusplus */

#endif 	/* __server_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

#ifndef __Iserver_INTERFACE_DEFINED__
#define __Iserver_INTERFACE_DEFINED__

/* interface Iserver */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_Iserver;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("426D34DC-2298-417E-A840-6CBDB5AE93E0")
    Iserver : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Method1( 
            /* [in] */ BSTR sPath,
            /* [retval][out] */ long __RPC_FAR *retVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Method2( 
            /* [in] */ long lin,
            /* [retval][out] */ long __RPC_FAR *retVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Method3( 
            /* [in] */ VARIANT vin,
            /* [retval][out] */ long __RPC_FAR *retVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Method4( 
            /* [in] */ BSTR sPath,
            /* [in] */ BSTR msg,
            /* [retval][out] */ long __RPC_FAR *retVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE HeapCorruption( 
            /* [in] */ BSTR strIn,
            /* [in] */ long bufSize,
            /* [retval][out] */ long __RPC_FAR *retVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IserverVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            Iserver __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            Iserver __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            Iserver __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            Iserver __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            Iserver __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            Iserver __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            Iserver __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Method1 )( 
            Iserver __RPC_FAR * This,
            /* [in] */ BSTR sPath,
            /* [retval][out] */ long __RPC_FAR *retVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Method2 )( 
            Iserver __RPC_FAR * This,
            /* [in] */ long lin,
            /* [retval][out] */ long __RPC_FAR *retVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Method3 )( 
            Iserver __RPC_FAR * This,
            /* [in] */ VARIANT vin,
            /* [retval][out] */ long __RPC_FAR *retVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Method4 )( 
            Iserver __RPC_FAR * This,
            /* [in] */ BSTR sPath,
            /* [in] */ BSTR msg,
            /* [retval][out] */ long __RPC_FAR *retVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *HeapCorruption )( 
            Iserver __RPC_FAR * This,
            /* [in] */ BSTR strIn,
            /* [in] */ long bufSize,
            /* [retval][out] */ long __RPC_FAR *retVal);
        
        END_INTERFACE
    } IserverVtbl;

    interface Iserver
    {
        CONST_VTBL struct IserverVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define Iserver_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define Iserver_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define Iserver_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define Iserver_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define Iserver_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define Iserver_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define Iserver_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define Iserver_Method1(This,sPath,retVal)	\
    (This)->lpVtbl -> Method1(This,sPath,retVal)

#define Iserver_Method2(This,lin,retVal)	\
    (This)->lpVtbl -> Method2(This,lin,retVal)

#define Iserver_Method3(This,vin,retVal)	\
    (This)->lpVtbl -> Method3(This,vin,retVal)

#define Iserver_Method4(This,sPath,msg,retVal)	\
    (This)->lpVtbl -> Method4(This,sPath,msg,retVal)

#define Iserver_HeapCorruption(This,strIn,bufSize,retVal)	\
    (This)->lpVtbl -> HeapCorruption(This,strIn,bufSize,retVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Iserver_Method1_Proxy( 
    Iserver __RPC_FAR * This,
    /* [in] */ BSTR sPath,
    /* [retval][out] */ long __RPC_FAR *retVal);


void __RPC_STUB Iserver_Method1_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Iserver_Method2_Proxy( 
    Iserver __RPC_FAR * This,
    /* [in] */ long lin,
    /* [retval][out] */ long __RPC_FAR *retVal);


void __RPC_STUB Iserver_Method2_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Iserver_Method3_Proxy( 
    Iserver __RPC_FAR * This,
    /* [in] */ VARIANT vin,
    /* [retval][out] */ long __RPC_FAR *retVal);


void __RPC_STUB Iserver_Method3_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Iserver_Method4_Proxy( 
    Iserver __RPC_FAR * This,
    /* [in] */ BSTR sPath,
    /* [in] */ BSTR msg,
    /* [retval][out] */ long __RPC_FAR *retVal);


void __RPC_STUB Iserver_Method4_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Iserver_HeapCorruption_Proxy( 
    Iserver __RPC_FAR * This,
    /* [in] */ BSTR strIn,
    /* [in] */ long bufSize,
    /* [retval][out] */ long __RPC_FAR *retVal);


void __RPC_STUB Iserver_HeapCorruption_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __Iserver_INTERFACE_DEFINED__ */



#ifndef __VULNLib_LIBRARY_DEFINED__
#define __VULNLib_LIBRARY_DEFINED__

/* library VULNLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_VULNLib;

#ifndef ___IserverEvents_DISPINTERFACE_DEFINED__
#define ___IserverEvents_DISPINTERFACE_DEFINED__

/* dispinterface _IserverEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__IserverEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("8712742F-B31D-4241-82D8-0B1C93D6F8A7")
    _IserverEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IserverEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            _IserverEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            _IserverEvents __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            _IserverEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            _IserverEvents __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            _IserverEvents __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            _IserverEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            _IserverEvents __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } _IserverEventsVtbl;

    interface _IserverEvents
    {
        CONST_VTBL struct _IserverEventsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IserverEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _IserverEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _IserverEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _IserverEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _IserverEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _IserverEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _IserverEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IserverEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_server;

#ifdef __cplusplus

class DECLSPEC_UUID("8EF2A07C-6E69-4144-96AA-2247D892A73D")
server;
#endif
#endif /* __VULNLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long __RPC_FAR *, unsigned long            , BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long __RPC_FAR *, BSTR __RPC_FAR * ); 

unsigned long             __RPC_USER  VARIANT_UserSize(     unsigned long __RPC_FAR *, unsigned long            , VARIANT __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  VARIANT_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, VARIANT __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  VARIANT_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, VARIANT __RPC_FAR * ); 
void                      __RPC_USER  VARIANT_UserFree(     unsigned long __RPC_FAR *, VARIANT __RPC_FAR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
