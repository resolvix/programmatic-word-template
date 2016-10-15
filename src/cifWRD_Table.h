// cifWRD_Table.h : Declaration of the cifWRD_Table

#ifndef __CIFWRD_TABLE_H_
#define __CIFWRD_TABLE_H_

#include "resource.h"       // main symbols

class cWRD_Table;

/////////////////////////////////////////////////////////////////////////////
// cifWRD_Table
class ATL_NO_VTABLE cifWRD_Table : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<cifWRD_Table, &CLSID_cifWRD_Table>,
	public IDispatchImpl<iifWRD_Table, &IID_iifWRD_Table, &LIBID_WordPRO>,
	public ISupportErrorInfoImpl< &IID_iifWRD_Table >
{
public:

DECLARE_REGISTRY_RESOURCEID(IDR_CIFWRD_TABLE)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(cifWRD_Table)
	COM_INTERFACE_ENTRY(iifWRD_Table)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// iifWRD_Table
public:

	cifWRD_Table(
			void
		);

	virtual ~cifWRD_Table(
			void
		);

	void create(
			cWRD_Table *p_pTable
		);

	STDMETHOD(setCell)(
			LONG p_lRowIndex,		// in
			BSTR p_bstrColumnId,	// in
			BSTR p_bstrValue		// in
		);

protected:

	cWRD_Table *m_pTable;
	
};

#endif //__CIFWRD_TABLE_H_
