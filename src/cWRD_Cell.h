// cWRD_Cell.h: interface for the cWRD_Cell class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CWRD_CELL_H__A47DCCF5_43E7_444F_AA0A_29A453610FAF__INCLUDED_)
#define AFX_CWRD_CELL_H__A47DCCF5_43E7_444F_AA0A_29A453610FAF__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class cWRD_Cell : public cWRD_Object
{
public:

	cWRD_Cell(
			LPDISPATCH p_pdCell
		);

	virtual ~cWRD_Cell();

	LONG getColumnIndex();

	virtual etType type( void );

protected:

	MSWord::Cell m_objCell;

};

#endif // !defined(AFX_CWRD_CELL_H__A47DCCF5_43E7_444F_AA0A_29A453610FAF__INCLUDED_)
