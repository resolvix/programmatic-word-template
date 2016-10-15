// cWRD_Row.h: interface for the cWRD_Row class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CWRD_ROW_H__642DE165_563B_490F_B234_2ACD83F0DE43__INCLUDED_)
#define AFX_CWRD_ROW_H__642DE165_563B_490F_B234_2ACD83F0DE43__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class cWRD_Row : public cWRD_Object
{
public:

	cWRD_Row(
			LPDISPATCH p_pdRow
		);

	virtual ~cWRD_Row();

	LONG cWRD_Row::getRowIndex( void );

	etType type( void )
	{
		return WORD_ROW;
	}

protected:

	MSWord::Row m_objRow;

};

#endif // !defined(AFX_CWRD_ROW_H__642DE165_563B_490F_B234_2ACD83F0DE43__INCLUDED_)
