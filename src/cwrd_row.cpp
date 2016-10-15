// cWRD_Row.cpp: implementation of the cWRD_Row class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"

namespace MSWord {

#include "msword9.h"

}

#include "cWRD_Object.h"
#include "cWRD_Row.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

/*****************************************************************************


 *****************************************************************************/

cWRD_Row::cWRD_Row(
		LPDISPATCH p_pdRow
	)
{
	p_pdRow->AddRef();

	m_objRow.AttachDispatch(
			p_pdRow, FALSE
		);

	return;
}

/*****************************************************************************


 *****************************************************************************/

cWRD_Row::~cWRD_Row()
{
	m_objRow.m_lpDispatch->Release();

	m_objRow.DetachDispatch();

	return;
}

/*****************************************************************************


 *****************************************************************************/

LONG cWRD_Row::getRowIndex( void )
{
	return m_objRow.GetIndex();
}

/*****************************************************************************

 End of Module

 *****************************************************************************/