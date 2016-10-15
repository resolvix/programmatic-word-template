// cWRD_Table.h: interface for the cWRD_Table class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CWRD_TABLE_H__78E8D261_F34D_4B5B_AFF8_D99043F89B4F__INCLUDED_)
#define AFX_CWRD_TABLE_H__78E8D261_F34D_4B5B_AFF8_D99043F89B4F__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class cWRD_Table : public cWRD_Object
{
public:

	cWRD_Table(
			std::string p_sId,
			LPDISPATCH p_pdTable
		);

	virtual ~cWRD_Table();	
	
	void defineBody(
			LONG p_lBodyRow_begin,
			LONG p_lBodyRow_end
		);

	void defineColumnId(
			std::string p_sId,
			LONG p_lIndex
		);

	void defineFooter(
			LONG p_lFooterRow_begin,
			LONG p_lFooterRow_end
		);

	void defineHeader(
			LONG p_lHeaderRow_begin,
			LONG p_lHeaderRow_end
		);

	void set(
			LONG p_lRowOffset,
			std::string p_sColumnId,
			std::string p_sValue
		);

	etType type( void )
	{
		return WORD_TABLE;
	}

protected:

	class emapColumn : public std::map< 
			std::string, 
			LONG, 
			std::less< std::string >, 
			std::allocator< LONG >
		> { /* */ };

	std::string m_sId;

	MSWord::Table m_objTable;

	emapColumn m_mapColumn;

	LONG m_lHeaderRow_begin;
	LONG m_lHeaderRow_end;

	LONG m_lBodyRow_begin;
	LONG m_lBodyRow_end;

	LONG m_lFooterRow_begin;
	LONG m_lFooterRow_end;

};

#endif // !defined(AFX_CWRD_TABLE_H__78E8D261_F34D_4B5B_AFF8_D99043F89B4F__INCLUDED_)
