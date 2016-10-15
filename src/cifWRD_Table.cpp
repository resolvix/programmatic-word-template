// cifWRD_Table.cpp : Implementation of cifWRD_Table
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

#include "cWRD_Object.h"
#include "cWRD_Table.h"
#include "cifWRD_Table.h"

/*****************************************************************************


 *****************************************************************************/

cifWRD_Table::cifWRD_Table(
		void  
	)
{
	m_pTable = NULL;

	return;
}

/*****************************************************************************


 *****************************************************************************/

cifWRD_Table::~cifWRD_Table(
		void		
	)
{
	m_pTable = NULL;

	return;
}

/*****************************************************************************


 *****************************************************************************/

void cifWRD_Table::create(
		cWRD_Table *p_pTable
	)
{
	m_pTable = p_pTable;

	return;
}

/*****************************************************************************


 *****************************************************************************/

STDMETHODIMP cifWRD_Table::setCell(
		LONG p_lRowIndex,		// in
		BSTR p_bstrColumnId,	// in
		BSTR p_bstrValue		// in
	)
{
	USES_CONVERSION;

	LPSTR lpstrColumnId = OLE2A(
			p_bstrColumnId
		);

	LPSTR lpstrValue = OLE2A(
			p_bstrValue
		);

	try
	{
		m_pTable->set( 
				p_lRowIndex,
				std::string(lpstrColumnId),
				std::string(lpstrValue)
			);	
	}
	catch(COleDispatchException *pE)
	{
		return cGEN_ErrorHandler::ComException(
				pE,
				IID_iifWRD_Table
			);
	}

	return S_OK;
}

/*****************************************************************************

 End of Module

 *****************************************************************************/