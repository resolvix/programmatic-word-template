/*****************************************************************************


 *****************************************************************************/

class cWRD_Object : public CObject
{
public:

	typedef enum 
	{
		WORD_TABLE,
		WORD_ROW,
		WORD_CELL
	}
	etType;

	virtual ~cWRD_Object() { return; }

	virtual etType type( void ) = 0;

};

/*****************************************************************************
	
 End of Module

 *****************************************************************************/