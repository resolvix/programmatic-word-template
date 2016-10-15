// cifWRD_Document.h : Declaration of the cifWRD_Document

#ifndef __CIFWRD_DOCUMENT_H_
#define __CIFWRD_DOCUMENT_H_

#include "resource.h"       // main symbols

class cWRD_Document;

/////////////////////////////////////////////////////////////////////////////
// cifWRD_Document
class ATL_NO_VTABLE cifWRD_Document : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<cifWRD_Document, &CLSID_cifWRD_Document>,
	public IDispatchImpl<iifWRD_Document, &IID_iifWRD_Document, &LIBID_WordPRO>,
	public ISupportErrorInfoImpl< &IID_iifWRD_Document >
{
public:

DECLARE_REGISTRY_RESOURCEID(IDR_CIFWRD_DOCUMENT)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(cifWRD_Document)
	COM_INTERFACE_ENTRY(iifWRD_Document)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// iifWRD_Document
public:

	cifWRD_Document( void );

	virtual ~cifWRD_Document();

	STDMETHOD(close)( void );

	STDMETHOD(create)(
			BSTR p_bstrTemplate
		);

	STDMETHOD(getTable)(
			BSTR p_bstrTableId,			// in
			iifWRD_Table **p_ppiTable	// out, retval
		);

	STDMETHOD(preprocess)( void );

	STDMETHOD(render)( void );

	STDMETHOD(set)( 
			BSTR p_bstrVarName,
			BSTR p_bstrValue
		);

	STDMETHOD(setTableCell)(
			BSTR p_bstrTableId,
			LONG p_lRowIndex,
			BSTR p_bstrColumnId,
			BSTR p_bstrValue
		);

	STDMETHOD(save)( void );

	STDMETHOD(saveas)( 
			BSTR p_bstrPathName	   
		);

protected:

	cWRD_Document *m_pDocument;

};

#endif //__CIFWRD_DOCUMENT_H_
