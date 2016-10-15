#include "stdafx.h"

#include <iostream>
#include <utility>
#include <string>
#include <map>
#include <vector>

#include "cGEN_Types.h"

#include "cGEN_MacroParser.h"

/*****************************************************************************


 *****************************************************************************/

BOOL cGEN_MacroParser::_seekMacro(
		std::string p_csText,
		LONG &p_rlCursor,
		std::string &p_rcsMacro,
		string_vector &p_rsvParam,
		LONG &p_rlPos_start,
		LONG &p_rlPos_end
	)
{
	//BOOL bResult;

	char chTemp;

	//char chBuffer[65536];
	char chBuffer2[1024];
	char chMacro[256];
	char chParam[256];

	string_vector svParam;
	std::string sResult;

	long lBuffer;
	long lBuffer2;
	long lMacro;
	long lParam;
	long lParam_count;

	long lPos_start;
	long lPos_end;

	long i;
	long i_max;

	etState state = S00;

	etToken token;

	i = p_rlCursor;
	i_max = p_csText.length();

	lBuffer = 0x00;
	
	while (i < i_max)
	{
		chTemp = p_csText[i];

		token = _tokenize(
				chTemp
			);

		switch (state)
		{
		case S00:

			switch (token)
			{
			case LEFT_BRACE:
				lPos_start = i;
				state = S01;
				break;

			}

			break;

		case S01:

			switch (token)
			{
			case DOLLAR_SIGN:
				state = S02;
				lMacro = 0x00;
				lBuffer2 = 0x00;
				break;

			default:
				state = S00;
				break;

			}
				
			break;

		case S02:

			switch (token)
			{
			case ALPHA:
			case DIGIT:
			case MINUS:
			case UNKNOWN:
				chMacro[ lMacro++ ] = chTemp;
				chBuffer2[ lBuffer2++ ] = chTemp;
				break;

			case LEFT_BRACKET:
				chMacro[ lMacro++ ] = 0x00;
				chBuffer2[ lBuffer2++ ] = chTemp;
				svParam.clear();
				lParam = 0x00;
				lParam_count = 0x00;
				state = S03;
				break;

			case RIGHT_BRACE:
				
				chBuffer2[ lBuffer2++ ] = 0x00;
				chMacro[ lMacro++ ] = 0x00;

				lPos_end = i;

				p_rlCursor = i+1;
				p_rcsMacro = chMacro;
				p_rsvParam = svParam;
				p_rlPos_start = lPos_start;
				p_rlPos_end = lPos_end;

				return TRUE;

				break;

			}

			break;

		case S03:

			chBuffer2[ lBuffer2++ ] = chTemp;

			switch (token)
			{
			case ALPHA:
			case DIGIT:
			case MINUS:
				chParam[ lParam++ ] = chTemp;
				break;

			case COMMA:
				chParam[ lParam++ ] = 0x00;
				svParam.insert(
						svParam.end(),
						chParam
					);
				lParam = 0x00;
				break;

			case RIGHT_BRACKET:
				chBuffer2[ lBuffer2++ ] = 0x00;
				chParam[ lParam++ ] = 0x00;
				svParam.insert(
						svParam.end(),
						chParam
					);

				state = S04;
				break;

			}

			break;

		case S04:

			switch (token)
			{
			case RIGHT_BRACE:

				lPos_end = i;
				p_rlCursor = i+1;
				p_rcsMacro = chMacro;
				p_rsvParam = svParam;

				p_rlPos_start = lPos_start;
				p_rlPos_end = lPos_end;

				return TRUE;

				break;

			}

			break;

		}

		i++;
	}

	return FALSE;
}

/******************************************************************************


 ******************************************************************************/
/*
std::string cGEN_MacroParser::_parse(
		std::string p_csText,
		etMode p_eMode,
		pftExecuteMacro p_pftExecuteMacro
	)
{
	BOOL bResult;

	char chTemp;

	char chBuffer[65536];
	char chBuffer2[1024];
	char chMacro[256];
	char chParam[256];

	string_vector svParam;
	std::string sResult;

	long lBuffer;
	long lBuffer2;
	long lMacro;
	long lParam;
	long lParam_count;

	long i;
	long i_max;

	etState state = S00;

	etToken token;

	i_max = p_csText.length();

	lBuffer = 0x00;
	
	for (i = 0x00; i < i_max; i++)
	{
		chTemp = p_csText[i];

		token = _tokenize(
				chTemp
			);

		switch (state)
		{
		case S00:

			switch (token)
			{
			case LEFT_BRACE:
				state = S01;
				break;

			default:
				chBuffer[lBuffer++] = chTemp;
				break;
			}

			break;

		case S01:

			switch (token)
			{
			case DOLLAR_SIGN:
				state = S02;
				lMacro = 0x00;
				lBuffer2 = 0x00;
				break;

			}
				
			break;

		case S02:

			switch (token)
			{
			case ALPHA:
			case DIGIT:
			case MINUS:
			case UNKNOWN:
				chMacro[ lMacro++ ] = chTemp;
				chBuffer2[ lBuffer2++ ] = chTemp;
				break;

			case LEFT_BRACKET:
				chMacro[ lMacro++ ] = 0x00;
				chBuffer2[ lBuffer2++ ] = chTemp;
				svParam.clear();
				lParam = 0x00;
				lParam_count = 0x00;
				state = S03;
				break;

			case RIGHT_BRACE:
				
				chBuffer2[ lBuffer2++ ] = 0x00;
				chMacro[ lMacro++ ] = 0x00;

				bResult = ((this-4)->*p_pftExecuteMacro)(
						p_eMode,
						chMacro,
						svParam,
						sResult
					);

				if (bResult == TRUE) // execute was successful
				{
					// append p_sValue to chBuffer
					LONG n;
					LONG n_max;

					n_max = sResult.length();

					for (n = 0x00; n < n_max; n++) {
						chBuffer[lBuffer++] = sResult[n];
					}
				}
				else
				{
					// just append as normal
					chBuffer[lBuffer++] = '{';
					chBuffer[lBuffer++] = '$';

					LONG n;
					LONG n_max;

					n_max = strlen(chBuffer2);

					for (n = 0x00; n < n_max; n++) {
						chBuffer[lBuffer++] = chBuffer2[n];
					}

					chBuffer[lBuffer++] = '}';

				}

				state = S00;

				break;

			}

			break;

		case S03:

			chBuffer2[ lBuffer2++ ] = chTemp;

			switch (token)
			{
			case ALPHA:
			case DIGIT:
			case MINUS:
				chParam[ lParam++ ] = chTemp;
				break;

			case COMMA:
				chParam[ lParam++ ] = 0x00;
				svParam.insert(
						svParam.end(),
						chParam
					);
				lParam = 0x00;
				break;

			case RIGHT_BRACKET:
				chBuffer2[ lBuffer2++ ] = 0x00;
				chParam[ lParam++ ] = 0x00;
				svParam.insert(
						svParam.end(),
						chParam
					);

				state = S04;

			}

			break;

		case S04:

			switch (token)
			{
			case RIGHT_BRACE:

				bResult = ((this - 4)->*p_pftExecuteMacro)(
						p_eMode,
						chMacro,
						svParam,
						sResult
					);

				if (bResult == TRUE) // execute was successful
				{
					// append p_sValue to chBuffer
					LONG n;
					LONG n_max;
					
					n_max = sResult.length();

					for (n = 0x00; n < n_max; n++) {
						chBuffer[lBuffer++] = sResult[n];
					}
				}
				else
				{
					// just append as normal
					chBuffer[lBuffer++] = '{';
					chBuffer[lBuffer++] = '$';

					LONG n;
					LONG n_max;

					n_max = strlen(chBuffer2);

					for (n = 0x00; n < n_max; n++) {
						chBuffer[lBuffer++] = chBuffer2[n];
					}

					chBuffer[lBuffer++] = '}';
				}

				state = S00;

				break;

			}

			break;

		}

	}

	chBuffer[lBuffer++] = 0x00;

	return std::string(chBuffer);
}*/

/*****************************************************************************


 *****************************************************************************/
 
cGEN_MacroParser::etToken cGEN_MacroParser::_tokenize(
		char p_Char
	)
{
	if ((p_Char >= 'A' && p_Char <= 'Z')
		|| (p_Char >= 'a' && p_Char <= 'z'))
		return ALPHA;

	if ((p_Char >= '0' && p_Char <= '9'))
		return DIGIT;

	switch (p_Char)
	{
	case ',':
		return COMMA; break;

	case '-':
		return MINUS; break;

	case '(':
		return LEFT_BRACKET; break;

	case ')':
		return RIGHT_BRACKET; break;

	case '{':
		return LEFT_BRACE; break;

	case '}':
		return RIGHT_BRACE; break;

	case '$':
		return DOLLAR_SIGN; break;
	}

	return UNKNOWN;		
}

/*****************************************************************************

 End of Module

 *****************************************************************************/