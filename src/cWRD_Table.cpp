// cWRD_Table.cpp: implementation of the cWRD_Table class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"

#include <iostream>
#include <string>
#include <map>

namespace MSWord {

#include "msword9.h"

}

#include "cWRD_Object.h"
#include "cWRD_Table.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

/*****************************************************************************


 *****************************************************************************/

cWRD_Table::cWRD_Table(
		std::string p_sId,
		LPDISPATCH p_pdTable
	)
{
	m_sId = p_sId;

	m_objTable.AttachDispatch(
			p_pdTable,
			FALSE
		);

	p_pdTable->AddRef();

	return;
}

/*****************************************************************************


 *****************************************************************************/

cWRD_Table::~cWRD_Table()
{
	m_mapColumn.clear();

	m_objTable.m_lpDispatch->Release();

	m_objTable.DetachDispatch();

	return;
}

/*****************************************************************************


 *****************************************************************************/

void cWRD_Table::defineBody(
		LONG p_lBodyRow_begin,
		LONG p_lBodyRow_end
	)
{
	m_lBodyRow_begin = p_lBodyRow_begin;
	m_lBodyRow_end = p_lBodyRow_end;

	return;
}

/*****************************************************************************


 *****************************************************************************/

void cWRD_Table::defineColumnId(
		std::string p_sId,
		LONG p_lIndex
	)
{
	m_mapColumn.insert(
			emapColumn::value_type(
					p_sId,
					p_lIndex
				)	
		);

	return;
}


/*****************************************************************************


 *****************************************************************************/

void cWRD_Table::defineFooter(
		LONG p_lFooterRow_begin,
		LONG p_lFooterRow_end
	)
{
	m_lFooterRow_begin = p_lFooterRow_begin;
	m_lFooterRow_end = p_lFooterRow_end;

	return;
}

/*****************************************************************************


 *****************************************************************************/

void cWRD_Table::defineHeader(
		LONG p_lHeaderRow_begin,
		LONG p_lHeaderRow_end
	)
{
	m_lHeaderRow_begin = p_lHeaderRow_begin;
	m_lHeaderRow_end = p_lHeaderRow_end;

	return;
}

/*****************************************************************************


 *****************************************************************************/

void cWRD_Table::set(
		LONG p_lRowOffset,
		std::string p_sColumnId,
		std::string p_sValue
	)
{
	LONG lRow;
	LONG lColumn;

	emapColumn::iterator itTemp;

	// get hold of table dispatch object
	MSWord::Rows objRows;
	MSWord::Row objRow;

	MSWord::Cells objCells;
	MSWord::Cell objCell;

	MSWord::Range objRange;

	lRow = m_lBodyRow_begin + p_lRowOffset + 1;

	itTemp = m_mapColumn.find(
			std::string(p_sColumnId)
		);

	if (itTemp != m_mapColumn.end())
	{
		lColumn = itTemp->second;

		if (lRow <= m_lBodyRow_end )
		{
			objRows = m_objTable.GetRows();

			objRow = objRows.Item(
					lRow
				);
			
			objCells = objRow.GetCells();

			objCell = objCells.Item(
					lColumn
				);

			objRange = objCell.GetRange();

			objRange.SetText(
					p_sValue.c_str()
				);
		}
	}

	return;
}

/*****************************************************************************
 
 End of Module

 *****************************************************************************/
