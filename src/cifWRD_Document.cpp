// cifWRD_Document.cpp : Implementation of cifWRD_Document
#include "stdafx.h"

#include <iostream>
#include <utility>
#include <string>
#include <map>
#include <vector>
#include <stack>

#include "WordPRO.h"

namespace MSWord {

#include "msword9.h"

}

#include "cGEN_Types.h"
#include "cGEN_ErrorHandler.h"
#include "cGEN_MacroParser.h"

#include "cWRD_Object.h"

#include "cWRD_stackObject.h"
#include "cWRD_Table.h"
#include "cWRD_setTable.h"
#include "cWRD_Document.h"
#include "cifWRD_Table.h"
#include "cifWRD_Document.h"

/*****************************************************************************

 *****************************************************************************/

cifWRD_Document::cifWRD_Document( void )
{
	m_pDocument = NULL;
}

/*****************************************************************************

 *****************************************************************************/

cifWRD_Document::~cifWRD_Document()
{
	if (m_pDocument != NULL)
	{
		close();
	}

	return;
}

/*****************************************************************************

 *****************************************************************************/

STDMETHODIMP cifWRD_Document::create(
		BSTR p_bstrTemplate
	)
{
	USES_CONVERSION;

	LPSTR lpstrTemplate = OLE2A(
			p_bstrTemplate
		);

	m_pDocument = new cWRD_Document();

	try
	{
		m_pDocument->create(
				lpstrTemplate
			);
	}
	catch(CException *pException)
	{
		return cGEN_ErrorHandler::ComException(
				pException,
				IID_iifWRD_Document
			);
	}

	return S_OK;
}


/*****************************************************************************

 *****************************************************************************/

STDMETHODIMP cifWRD_Document::close( 
		void
	)
{
	try
	{
		m_pDocument->close();

		delete m_pDocument;

		m_pDocument = NULL;
	}
	catch(COleDispatchException *pE)
	{
		return cGEN_ErrorHandler::ComException(
				pE,
				IID_iifWRD_Document
			);
	}

	return S_OK;
}

/*****************************************************************************

 *****************************************************************************/

STDMETHODIMP cifWRD_Document::getTable(
		BSTR p_bstrTableId,				// in
		iifWRD_Table **p_ppiTable		// out, retval
	)
{
	USES_CONVERSION;

	LPSTR lpstrTableId = OLE2A(
			p_bstrTableId
		);

	cWRD_Table *pTable;

	CComObject<cifWRD_Table> *pdTable;

	try
	{

		CComObject<cifWRD_Table>::CreateInstance(
				&pdTable
			);

		pdTable->AddRef();


		pTable = m_pDocument->getTable(
				lpstrTableId
			);
		
		if (pTable != NULL)
		{
			pdTable->create(
					pTable
				);

			*p_ppiTable = pdTable; 

			return S_OK;
		}

	}
	catch(COleDispatchException *pE)
	{
		return cGEN_ErrorHandler::ComException(
				pE,
				IID_iifWRD_Document
			);
	}

	return S_OK;
}

/*****************************************************************************

 *****************************************************************************/

STDMETHODIMP cifWRD_Document::preprocess( void )
{
	try
	{
		m_pDocument->preprocess();
	}
	catch(COleDispatchException *pE)
	{
		return cGEN_ErrorHandler::ComException(
				pE,
				IID_iifWRD_Document
			);
	}

	return S_OK;
}


/*****************************************************************************

 *****************************************************************************/

STDMETHODIMP cifWRD_Document::render( void )
{
	try
	{
		m_pDocument->render();
	}
	catch(COleDispatchException *pE)
	{
		return cGEN_ErrorHandler::ComException(
				pE,
				IID_iifWRD_Document
			);
	}

	return S_OK;
}

/*****************************************************************************

 *****************************************************************************/

STDMETHODIMP cifWRD_Document::set( 
		BSTR p_bstrVarName,
		BSTR p_bstrValue
	)
{
	USES_CONVERSION;

	LPSTR lpstrVarName = OLE2A(
			p_bstrVarName
		);

	LPSTR lpstrValue = OLE2A(
			p_bstrValue
		);

	try
	{
		m_pDocument->set(
				lpstrVarName,
				lpstrValue
			);
	}
	catch(COleDispatchException *pE)
	{
		return cGEN_ErrorHandler::ComException(
				pE,
				IID_iifWRD_Document
			);
	}

	return S_OK;
}

/*****************************************************************************

 *****************************************************************************/

STDMETHODIMP cifWRD_Document::setTableCell(
		BSTR p_bstrTableId,
		LONG p_lRowIndex,
		BSTR p_bstrColumnId,
		BSTR p_bstrValue
	)
{
	USES_CONVERSION;

	LPSTR lpstrTableId = OLE2A(
			p_bstrTableId
		);

	LPSTR lpstrColumnId = OLE2A(
			p_bstrColumnId
		);

	LPSTR lpstrValue = OLE2A(
			p_bstrValue
		);

	try
	{
		m_pDocument->setTableCell(
				lpstrTableId,
				p_lRowIndex,
				lpstrColumnId,
				lpstrValue
			);
	}
	catch(COleDispatchException *pE)
	{
		return cGEN_ErrorHandler::ComException(
				pE,
				IID_iifWRD_Document
			);
	}

	return S_OK;
}

/*****************************************************************************

 *****************************************************************************/

HRESULT cifWRD_Document::save( void )
{
	try
	{
		m_pDocument->save();
	}
	catch(COleDispatchException *pE)
	{
		return cGEN_ErrorHandler::ComException(
				pE,
				IID_iifWRD_Document
			);
	}

	return S_OK;
}

/*****************************************************************************

 *****************************************************************************/

HRESULT cifWRD_Document::saveas( 
		BSTR p_bstrPathName	   
	)
{
	USES_CONVERSION;

	LPSTR lpstrPathName = OLE2A(
			p_bstrPathName
		);

	try
	{
		m_pDocument->saveas( 
				lpstrPathName
			);
	}
	catch(COleDispatchException *pE)
	{
		return cGEN_ErrorHandler::ComException(
				pE,
				IID_iifWRD_Document
			);
	}

	return S_OK;
}

/*****************************************************************************

 *****************************************************************************/