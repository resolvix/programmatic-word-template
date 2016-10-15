
#pragma warning( disable: 4049 )  /* more than 64k source lines */

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 6.00.0347 */
/* at Wed Feb 28 16:14:52 2007
 */
/* Compiler settings for Wordpro.odl:
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

#ifndef __WordPRO_i_h__
#define __WordPRO_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __iifWRD_Table_FWD_DEFINED__
#define __iifWRD_Table_FWD_DEFINED__
typedef interface iifWRD_Table iifWRD_Table;
#endif 	/* __iifWRD_Table_FWD_DEFINED__ */


#ifndef __iifWRD_Document_FWD_DEFINED__
#define __iifWRD_Document_FWD_DEFINED__
typedef interface iifWRD_Document iifWRD_Document;
#endif 	/* __iifWRD_Document_FWD_DEFINED__ */


#ifndef __cifWRD_Document_FWD_DEFINED__
#define __cifWRD_Document_FWD_DEFINED__

#ifdef __cplusplus
typedef class cifWRD_Document cifWRD_Document;
#else
typedef struct cifWRD_Document cifWRD_Document;
#endif /* __cplusplus */

#endif 	/* __cifWRD_Document_FWD_DEFINED__ */


#ifndef __cifWRD_Table_FWD_DEFINED__
#define __cifWRD_Table_FWD_DEFINED__

#ifdef __cplusplus
typedef class cifWRD_Table cifWRD_Table;
#else
typedef struct cifWRD_Table cifWRD_Table;
#endif /* __cplusplus */

#endif 	/* __cifWRD_Table_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"

#ifdef __cplusplus
extern "C"{
#endif 

void * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void * ); 


#ifndef __WordPRO_LIBRARY_DEFINED__
#define __WordPRO_LIBRARY_DEFINED__

/* library WordPRO */
/* [version][uuid] */ 




EXTERN_C const IID LIBID_WordPRO;

#ifndef __iifWRD_Table_INTERFACE_DEFINED__
#define __iifWRD_Table_INTERFACE_DEFINED__

/* interface iifWRD_Table */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_iifWRD_Table;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("5061A8BD-1850-49E3-88AF-962467535C62")
    iifWRD_Table : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE setCell( 
            /* [in] */ LONG p_lRowIndex,
            /* [in] */ BSTR p_bstrColumnId,
            /* [in] */ BSTR p_bstrValue) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct iifWRD_TableVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            iifWRD_Table * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            iifWRD_Table * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            iifWRD_Table * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            iifWRD_Table * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            iifWRD_Table * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            iifWRD_Table * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            iifWRD_Table * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *setCell )( 
            iifWRD_Table * This,
            /* [in] */ LONG p_lRowIndex,
            /* [in] */ BSTR p_bstrColumnId,
            /* [in] */ BSTR p_bstrValue);
        
        END_INTERFACE
    } iifWRD_TableVtbl;

    interface iifWRD_Table
    {
        CONST_VTBL struct iifWRD_TableVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define iifWRD_Table_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define iifWRD_Table_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define iifWRD_Table_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define iifWRD_Table_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define iifWRD_Table_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define iifWRD_Table_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define iifWRD_Table_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define iifWRD_Table_setCell(This,p_lRowIndex,p_bstrColumnId,p_bstrValue)	\
    (This)->lpVtbl -> setCell(This,p_lRowIndex,p_bstrColumnId,p_bstrValue)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE iifWRD_Table_setCell_Proxy( 
    iifWRD_Table * This,
    /* [in] */ LONG p_lRowIndex,
    /* [in] */ BSTR p_bstrColumnId,
    /* [in] */ BSTR p_bstrValue);


void __RPC_STUB iifWRD_Table_setCell_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __iifWRD_Table_INTERFACE_DEFINED__ */


#ifndef __iifWRD_Document_INTERFACE_DEFINED__
#define __iifWRD_Document_INTERFACE_DEFINED__

/* interface iifWRD_Document */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_iifWRD_Document;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("6B47EE8C-E208-4B4A-91AA-9123E558A27F")
    iifWRD_Document : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE close( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE create( 
            /* [in] */ BSTR p_bstrTemplate) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE getTable( 
            /* [in] */ BSTR p_bstrTableId,
            /* [retval][out] */ iifWRD_Table **p_ppiTable) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE preprocess( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE render( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE set( 
            /* [in] */ BSTR p_bstrVarName,
            /* [in] */ BSTR p_bstrValue) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE setTableCell( 
            /* [in] */ BSTR p_bstrTableId,
            /* [in] */ LONG p_lRowIndex,
            /* [in] */ BSTR p_bstrColumnId,
            /* [in] */ BSTR p_bstrValue) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE save( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE saveas( 
            /* [in] */ BSTR p_bstrPathName) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct iifWRD_DocumentVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            iifWRD_Document * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            iifWRD_Document * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            iifWRD_Document * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            iifWRD_Document * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            iifWRD_Document * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            iifWRD_Document * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            iifWRD_Document * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *close )( 
            iifWRD_Document * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *create )( 
            iifWRD_Document * This,
            /* [in] */ BSTR p_bstrTemplate);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *getTable )( 
            iifWRD_Document * This,
            /* [in] */ BSTR p_bstrTableId,
            /* [retval][out] */ iifWRD_Table **p_ppiTable);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *preprocess )( 
            iifWRD_Document * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *render )( 
            iifWRD_Document * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *set )( 
            iifWRD_Document * This,
            /* [in] */ BSTR p_bstrVarName,
            /* [in] */ BSTR p_bstrValue);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *setTableCell )( 
            iifWRD_Document * This,
            /* [in] */ BSTR p_bstrTableId,
            /* [in] */ LONG p_lRowIndex,
            /* [in] */ BSTR p_bstrColumnId,
            /* [in] */ BSTR p_bstrValue);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *save )( 
            iifWRD_Document * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *saveas )( 
            iifWRD_Document * This,
            /* [in] */ BSTR p_bstrPathName);
        
        END_INTERFACE
    } iifWRD_DocumentVtbl;

    interface iifWRD_Document
    {
        CONST_VTBL struct iifWRD_DocumentVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define iifWRD_Document_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define iifWRD_Document_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define iifWRD_Document_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define iifWRD_Document_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define iifWRD_Document_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define iifWRD_Document_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define iifWRD_Document_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define iifWRD_Document_close(This)	\
    (This)->lpVtbl -> close(This)

#define iifWRD_Document_create(This,p_bstrTemplate)	\
    (This)->lpVtbl -> create(This,p_bstrTemplate)

#define iifWRD_Document_getTable(This,p_bstrTableId,p_ppiTable)	\
    (This)->lpVtbl -> getTable(This,p_bstrTableId,p_ppiTable)

#define iifWRD_Document_preprocess(This)	\
    (This)->lpVtbl -> preprocess(This)

#define iifWRD_Document_render(This)	\
    (This)->lpVtbl -> render(This)

#define iifWRD_Document_set(This,p_bstrVarName,p_bstrValue)	\
    (This)->lpVtbl -> set(This,p_bstrVarName,p_bstrValue)

#define iifWRD_Document_setTableCell(This,p_bstrTableId,p_lRowIndex,p_bstrColumnId,p_bstrValue)	\
    (This)->lpVtbl -> setTableCell(This,p_bstrTableId,p_lRowIndex,p_bstrColumnId,p_bstrValue)

#define iifWRD_Document_save(This)	\
    (This)->lpVtbl -> save(This)

#define iifWRD_Document_saveas(This,p_bstrPathName)	\
    (This)->lpVtbl -> saveas(This,p_bstrPathName)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE iifWRD_Document_close_Proxy( 
    iifWRD_Document * This);


void __RPC_STUB iifWRD_Document_close_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE iifWRD_Document_create_Proxy( 
    iifWRD_Document * This,
    /* [in] */ BSTR p_bstrTemplate);


void __RPC_STUB iifWRD_Document_create_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE iifWRD_Document_getTable_Proxy( 
    iifWRD_Document * This,
    /* [in] */ BSTR p_bstrTableId,
    /* [retval][out] */ iifWRD_Table **p_ppiTable);


void __RPC_STUB iifWRD_Document_getTable_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE iifWRD_Document_preprocess_Proxy( 
    iifWRD_Document * This);


void __RPC_STUB iifWRD_Document_preprocess_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE iifWRD_Document_render_Proxy( 
    iifWRD_Document * This);


void __RPC_STUB iifWRD_Document_render_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE iifWRD_Document_set_Proxy( 
    iifWRD_Document * This,
    /* [in] */ BSTR p_bstrVarName,
    /* [in] */ BSTR p_bstrValue);


void __RPC_STUB iifWRD_Document_set_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE iifWRD_Document_setTableCell_Proxy( 
    iifWRD_Document * This,
    /* [in] */ BSTR p_bstrTableId,
    /* [in] */ LONG p_lRowIndex,
    /* [in] */ BSTR p_bstrColumnId,
    /* [in] */ BSTR p_bstrValue);


void __RPC_STUB iifWRD_Document_setTableCell_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE iifWRD_Document_save_Proxy( 
    iifWRD_Document * This);


void __RPC_STUB iifWRD_Document_save_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE iifWRD_Document_saveas_Proxy( 
    iifWRD_Document * This,
    /* [in] */ BSTR p_bstrPathName);


void __RPC_STUB iifWRD_Document_saveas_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __iifWRD_Document_INTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_cifWRD_Document;

#ifdef __cplusplus

class DECLSPEC_UUID("DE108E32-4199-4B82-B347-975E832A3535")
cifWRD_Document;
#endif

EXTERN_C const CLSID CLSID_cifWRD_Table;

#ifdef __cplusplus

class DECLSPEC_UUID("310083A8-4ED0-4D51-ABF7-8175B6F4AAA5")
cifWRD_Table;
#endif
#endif /* __WordPRO_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


