
#pragma warning( disable: 4049 )  /* more than 64k source lines */

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 6.00.0347 */
/* at Thu Sep 26 01:57:54 2002
 */
/* Compiler settings for Common.idl:
    Oicf, W1, Zp8, env=Win32 (32b run)
    protocol : dce , ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __Common_h__
#define __Common_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IConfig_FWD_DEFINED__
#define __IConfig_FWD_DEFINED__
typedef interface IConfig IConfig;
#endif 	/* __IConfig_FWD_DEFINED__ */


#ifndef __IDB_FWD_DEFINED__
#define __IDB_FWD_DEFINED__
typedef interface IDB IDB;
#endif 	/* __IDB_FWD_DEFINED__ */


#ifndef __Config_FWD_DEFINED__
#define __Config_FWD_DEFINED__

#ifdef __cplusplus
typedef class Config Config;
#else
typedef struct Config Config;
#endif /* __cplusplus */

#endif 	/* __Config_FWD_DEFINED__ */


#ifndef __DB_FWD_DEFINED__
#define __DB_FWD_DEFINED__

#ifdef __cplusplus
typedef class DB DB;
#else
typedef struct DB DB;
#endif /* __cplusplus */

#endif 	/* __DB_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 

void * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void * ); 


#ifndef __AU_CommonLib_LIBRARY_DEFINED__
#define __AU_CommonLib_LIBRARY_DEFINED__

/* library AU_CommonLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_AU_CommonLib;

#ifndef __IConfig_INTERFACE_DEFINED__
#define __IConfig_INTERFACE_DEFINED__

/* interface IConfig */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IConfig;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("C3BC9E26-E130-43CA-906D-25D9CB706597")
    IConfig : public IDispatch
    {
    public:
    };
    
#else 	/* C style interface */

    typedef struct IConfigVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IConfig * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IConfig * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IConfig * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IConfig * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IConfig * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IConfig * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IConfig * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } IConfigVtbl;

    interface IConfig
    {
        CONST_VTBL struct IConfigVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IConfig_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IConfig_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IConfig_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IConfig_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IConfig_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IConfig_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IConfig_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IConfig_INTERFACE_DEFINED__ */


#ifndef __IDB_INTERFACE_DEFINED__
#define __IDB_INTERFACE_DEFINED__

/* interface IDB */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IDB;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("FDA8DEE3-9B92-4CD5-B4D7-35EFC10A3F6E")
    IDB : public IDispatch
    {
    public:
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_ConnectionKey( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_ConnectionKey( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_ConnectionValue( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_ConnectionValue( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_CursorType( 
            /* [retval][out] */ int *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_CursorType( 
            /* [in] */ int newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_LockType( 
            /* [retval][out] */ int *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_LockType( 
            /* [in] */ int newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Timeout( 
            /* [retval][out] */ long *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_Timeout( 
            /* [in] */ long newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Connection( 
            /* [retval][out] */ /* external definition not present */ _Connection **pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Command( 
            /* [retval][out] */ /* external definition not present */ _Command **pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_IsOpen( 
            /* [retval][out] */ VARIANT_BOOL *pVal) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Open( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Close( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE LockDB( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE UnlockDB( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE NewCommand( 
            /* [retval][out] */ /* external definition not present */ _Command **pVal) = 0;
        
        virtual /* [id][vararg] */ HRESULT STDMETHODCALLTYPE Execute( 
            /* [in] */ BSTR sSQL,
            /* [in] */ SAFEARRAY * *aVals,
            /* [retval][out] */ /* external definition not present */ _Recordset **pRet) = 0;
        
        virtual /* [id][vararg] */ HRESULT STDMETHODCALLTYPE ExecuteNM( 
            /* [in] */ BSTR sSQL,
            /* [in] */ BSTR sParms,
            /* [in] */ SAFEARRAY * *aVals,
            /* [retval][out] */ /* external definition not present */ _Recordset **pRet) = 0;
        
        virtual /* [id][vararg] */ HRESULT STDMETHODCALLTYPE ExecuteVal( 
            /* [in] */ BSTR sSQL,
            /* [in] */ SAFEARRAY * *aVals,
            /* [retval][out] */ VARIANT *pRet) = 0;
        
        virtual /* [id][vararg] */ HRESULT STDMETHODCALLTYPE ExecuteNR( 
            /* [in] */ BSTR sSQL,
            /* [in] */ SAFEARRAY * *aVals) = 0;
        
        virtual /* [id][vararg] */ HRESULT STDMETHODCALLTYPE ExecuteNMNR( 
            /* [in] */ BSTR sSQL,
            /* [in] */ BSTR sParms,
            /* [in] */ SAFEARRAY * *aVals) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IDBVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IDB * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IDB * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IDB * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IDB * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IDB * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IDB * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IDB * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ConnectionKey )( 
            IDB * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_ConnectionKey )( 
            IDB * This,
            /* [in] */ BSTR newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ConnectionValue )( 
            IDB * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_ConnectionValue )( 
            IDB * This,
            /* [in] */ BSTR newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_CursorType )( 
            IDB * This,
            /* [retval][out] */ int *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_CursorType )( 
            IDB * This,
            /* [in] */ int newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_LockType )( 
            IDB * This,
            /* [retval][out] */ int *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_LockType )( 
            IDB * This,
            /* [in] */ int newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Timeout )( 
            IDB * This,
            /* [retval][out] */ long *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Timeout )( 
            IDB * This,
            /* [in] */ long newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Connection )( 
            IDB * This,
            /* [retval][out] */ /* external definition not present */ _Connection **pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Command )( 
            IDB * This,
            /* [retval][out] */ /* external definition not present */ _Command **pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_IsOpen )( 
            IDB * This,
            /* [retval][out] */ VARIANT_BOOL *pVal);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Open )( 
            IDB * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Close )( 
            IDB * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *LockDB )( 
            IDB * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *UnlockDB )( 
            IDB * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *NewCommand )( 
            IDB * This,
            /* [retval][out] */ /* external definition not present */ _Command **pVal);
        
        /* [id][vararg] */ HRESULT ( STDMETHODCALLTYPE *Execute )( 
            IDB * This,
            /* [in] */ BSTR sSQL,
            /* [in] */ SAFEARRAY * *aVals,
            /* [retval][out] */ /* external definition not present */ _Recordset **pRet);
        
        /* [id][vararg] */ HRESULT ( STDMETHODCALLTYPE *ExecuteNM )( 
            IDB * This,
            /* [in] */ BSTR sSQL,
            /* [in] */ BSTR sParms,
            /* [in] */ SAFEARRAY * *aVals,
            /* [retval][out] */ /* external definition not present */ _Recordset **pRet);
        
        /* [id][vararg] */ HRESULT ( STDMETHODCALLTYPE *ExecuteVal )( 
            IDB * This,
            /* [in] */ BSTR sSQL,
            /* [in] */ SAFEARRAY * *aVals,
            /* [retval][out] */ VARIANT *pRet);
        
        /* [id][vararg] */ HRESULT ( STDMETHODCALLTYPE *ExecuteNR )( 
            IDB * This,
            /* [in] */ BSTR sSQL,
            /* [in] */ SAFEARRAY * *aVals);
        
        /* [id][vararg] */ HRESULT ( STDMETHODCALLTYPE *ExecuteNMNR )( 
            IDB * This,
            /* [in] */ BSTR sSQL,
            /* [in] */ BSTR sParms,
            /* [in] */ SAFEARRAY * *aVals);
        
        END_INTERFACE
    } IDBVtbl;

    interface IDB
    {
        CONST_VTBL struct IDBVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IDB_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IDB_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IDB_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IDB_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IDB_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IDB_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IDB_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IDB_get_ConnectionKey(This,pVal)	\
    (This)->lpVtbl -> get_ConnectionKey(This,pVal)

#define IDB_put_ConnectionKey(This,newVal)	\
    (This)->lpVtbl -> put_ConnectionKey(This,newVal)

#define IDB_get_ConnectionValue(This,pVal)	\
    (This)->lpVtbl -> get_ConnectionValue(This,pVal)

#define IDB_put_ConnectionValue(This,newVal)	\
    (This)->lpVtbl -> put_ConnectionValue(This,newVal)

#define IDB_get_CursorType(This,pVal)	\
    (This)->lpVtbl -> get_CursorType(This,pVal)

#define IDB_put_CursorType(This,newVal)	\
    (This)->lpVtbl -> put_CursorType(This,newVal)

#define IDB_get_LockType(This,pVal)	\
    (This)->lpVtbl -> get_LockType(This,pVal)

#define IDB_put_LockType(This,newVal)	\
    (This)->lpVtbl -> put_LockType(This,newVal)

#define IDB_get_Timeout(This,pVal)	\
    (This)->lpVtbl -> get_Timeout(This,pVal)

#define IDB_put_Timeout(This,newVal)	\
    (This)->lpVtbl -> put_Timeout(This,newVal)

#define IDB_get_Connection(This,pVal)	\
    (This)->lpVtbl -> get_Connection(This,pVal)

#define IDB_get_Command(This,pVal)	\
    (This)->lpVtbl -> get_Command(This,pVal)

#define IDB_get_IsOpen(This,pVal)	\
    (This)->lpVtbl -> get_IsOpen(This,pVal)

#define IDB_Open(This)	\
    (This)->lpVtbl -> Open(This)

#define IDB_Close(This)	\
    (This)->lpVtbl -> Close(This)

#define IDB_LockDB(This)	\
    (This)->lpVtbl -> LockDB(This)

#define IDB_UnlockDB(This)	\
    (This)->lpVtbl -> UnlockDB(This)

#define IDB_NewCommand(This,pVal)	\
    (This)->lpVtbl -> NewCommand(This,pVal)

#define IDB_Execute(This,sSQL,aVals,pRet)	\
    (This)->lpVtbl -> Execute(This,sSQL,aVals,pRet)

#define IDB_ExecuteNM(This,sSQL,sParms,aVals,pRet)	\
    (This)->lpVtbl -> ExecuteNM(This,sSQL,sParms,aVals,pRet)

#define IDB_ExecuteVal(This,sSQL,aVals,pRet)	\
    (This)->lpVtbl -> ExecuteVal(This,sSQL,aVals,pRet)

#define IDB_ExecuteNR(This,sSQL,aVals)	\
    (This)->lpVtbl -> ExecuteNR(This,sSQL,aVals)

#define IDB_ExecuteNMNR(This,sSQL,sParms,aVals)	\
    (This)->lpVtbl -> ExecuteNMNR(This,sSQL,sParms,aVals)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [id][propget] */ HRESULT STDMETHODCALLTYPE IDB_get_ConnectionKey_Proxy( 
    IDB * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB IDB_get_ConnectionKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE IDB_put_ConnectionKey_Proxy( 
    IDB * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IDB_put_ConnectionKey_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE IDB_get_ConnectionValue_Proxy( 
    IDB * This,
    /* [retval][out] */ BSTR *pVal);


void __RPC_STUB IDB_get_ConnectionValue_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE IDB_put_ConnectionValue_Proxy( 
    IDB * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IDB_put_ConnectionValue_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE IDB_get_CursorType_Proxy( 
    IDB * This,
    /* [retval][out] */ int *pVal);


void __RPC_STUB IDB_get_CursorType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE IDB_put_CursorType_Proxy( 
    IDB * This,
    /* [in] */ int newVal);


void __RPC_STUB IDB_put_CursorType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE IDB_get_LockType_Proxy( 
    IDB * This,
    /* [retval][out] */ int *pVal);


void __RPC_STUB IDB_get_LockType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE IDB_put_LockType_Proxy( 
    IDB * This,
    /* [in] */ int newVal);


void __RPC_STUB IDB_put_LockType_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE IDB_get_Timeout_Proxy( 
    IDB * This,
    /* [retval][out] */ long *pVal);


void __RPC_STUB IDB_get_Timeout_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propput] */ HRESULT STDMETHODCALLTYPE IDB_put_Timeout_Proxy( 
    IDB * This,
    /* [in] */ long newVal);


void __RPC_STUB IDB_put_Timeout_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE IDB_get_Connection_Proxy( 
    IDB * This,
    /* [retval][out] */ /* external definition not present */ _Connection **pVal);


void __RPC_STUB IDB_get_Connection_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE IDB_get_Command_Proxy( 
    IDB * This,
    /* [retval][out] */ /* external definition not present */ _Command **pVal);


void __RPC_STUB IDB_get_Command_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][propget] */ HRESULT STDMETHODCALLTYPE IDB_get_IsOpen_Proxy( 
    IDB * This,
    /* [retval][out] */ VARIANT_BOOL *pVal);


void __RPC_STUB IDB_get_IsOpen_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IDB_Open_Proxy( 
    IDB * This);


void __RPC_STUB IDB_Open_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IDB_Close_Proxy( 
    IDB * This);


void __RPC_STUB IDB_Close_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IDB_LockDB_Proxy( 
    IDB * This);


void __RPC_STUB IDB_LockDB_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IDB_UnlockDB_Proxy( 
    IDB * This);


void __RPC_STUB IDB_UnlockDB_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IDB_NewCommand_Proxy( 
    IDB * This,
    /* [retval][out] */ /* external definition not present */ _Command **pVal);


void __RPC_STUB IDB_NewCommand_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][vararg] */ HRESULT STDMETHODCALLTYPE IDB_Execute_Proxy( 
    IDB * This,
    /* [in] */ BSTR sSQL,
    /* [in] */ SAFEARRAY * *aVals,
    /* [retval][out] */ /* external definition not present */ _Recordset **pRet);


void __RPC_STUB IDB_Execute_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][vararg] */ HRESULT STDMETHODCALLTYPE IDB_ExecuteNM_Proxy( 
    IDB * This,
    /* [in] */ BSTR sSQL,
    /* [in] */ BSTR sParms,
    /* [in] */ SAFEARRAY * *aVals,
    /* [retval][out] */ /* external definition not present */ _Recordset **pRet);


void __RPC_STUB IDB_ExecuteNM_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][vararg] */ HRESULT STDMETHODCALLTYPE IDB_ExecuteVal_Proxy( 
    IDB * This,
    /* [in] */ BSTR sSQL,
    /* [in] */ SAFEARRAY * *aVals,
    /* [retval][out] */ VARIANT *pRet);


void __RPC_STUB IDB_ExecuteVal_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][vararg] */ HRESULT STDMETHODCALLTYPE IDB_ExecuteNR_Proxy( 
    IDB * This,
    /* [in] */ BSTR sSQL,
    /* [in] */ SAFEARRAY * *aVals);


void __RPC_STUB IDB_ExecuteNR_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id][vararg] */ HRESULT STDMETHODCALLTYPE IDB_ExecuteNMNR_Proxy( 
    IDB * This,
    /* [in] */ BSTR sSQL,
    /* [in] */ BSTR sParms,
    /* [in] */ SAFEARRAY * *aVals);


void __RPC_STUB IDB_ExecuteNMNR_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IDB_INTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_Config;

#ifdef __cplusplus

class DECLSPEC_UUID("18B5A39D-092D-4916-87A5-CF627A34455B")
Config;
#endif

EXTERN_C const CLSID CLSID_DB;

#ifdef __cplusplus

class DECLSPEC_UUID("D406FC81-9412-48BA-BCCD-12EADDBB073A")
DB;
#endif
#endif /* __AU_CommonLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


