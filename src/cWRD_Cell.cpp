// cWRD_Cell.cpp: implementation of the cWRD_Cell class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"

namespace MSWord {

#include "msword9.h"

}

#include "cWRD_Object.h"
#include "cWRD_Cell.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

/*****************************************************************************


 *****************************************************************************/

cWRD_Cell::cWRD_Cell(
		LPDISPATCH p_pdCell	 
	)
{
	m_objCell.AttachDispatch(
			p_pdCell,
			FALSE
		);

	m_objCell.m_lpDispatch->AddRef();

	return;
}

/*****************************************************************************


 *****************************************************************************/

cWRD_Cell::~cWRD_Cell()
{
	m_objCell.m_lpDispatch->Release();

	m_objCell.DetachDispatch();

	return;
}

/*****************************************************************************


 *****************************************************************************/

LONG cWRD_Cell::getColumnIndex()
{
	return m_objCell.GetColumnIndex();
}

/*****************************************************************************


 *****************************************************************************/

cWRD_Cell::etType cWRD_Cell::type( void )
{
	return etType::WORD_CELL;
}

/*****************************************************************************

 End of Module

 *****************************************************************************/