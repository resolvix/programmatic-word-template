

class cGEN_MacroParser
{
protected:
	
	typedef enum etMode
	{
		modeTable,
		modeParagraph
	};

	typedef enum etState
	{
		S00, S01, S02, S03, S04
	};

	typedef enum etToken
	{
		ALPHA,
		DIGIT,
		COMMA,
		MINUS,
		LEFT_BRACKET,
		RIGHT_BRACKET,
		LEFT_BRACE,
		RIGHT_BRACE,
		DOLLAR_SIGN,
		UNKNOWN
	};

	etToken _tokenize(
			char p_Char
		);

/*	std::string _parse(
			std::string p_csText,
			etMode p_eMode,
			pftExecuteMacro p_pftExecuteMacro
		);*/

	BOOL _seekMacro(
			std::string p_csText,
			LONG &p_rlCursor,
			std::string &p_rcsMacro,
			string_vector &p_rsvParam,
			LONG &p_rlPos_start,
			LONG &p_rlPos_end
		);

};