// cWRD_Document.cpp: implementation of the cWRD_Document class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"

#include <iostream>
#include <utility>
#include <string>
#include <map>
#include <vector>
#include <stack>

namespace MSWord {

#include "msword9.h"

}

#include "cGEN_Types.h"
#include "cGEN_MacroParser.h"

#include "cWRD_Object.h"
#include "cWRD_stackObject.h"

#include "cWRD_Cell.h"

#include "cWRD_Row.h"

#include "cWRD_Table.h"
#include "cWRD_setTable.h"

#include "cWRD_Document.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

/*****************************************************************************



 *****************************************************************************/

cWRD_Document::cWRD_Document()
{
	return;
}

/*****************************************************************************


 *****************************************************************************/

cWRD_Document::~cWRD_Document()
{
	cWRD_Table *pTable;

	cWRD_mapTable::iterator itTemp;

	itTemp = m_mapTable.begin();

	while (itTemp != m_mapTable.end())
	{
		pTable = itTemp->second;

		delete[] pTable;

		itTemp++;
	}

	return;
}

/*****************************************************************************


 *****************************************************************************/

void cWRD_Document::create(
		char *p_sTemplate			 			 
	)
{	
	COleVariant vsTemplate(
			p_sTemplate
		);

	COleVariant vsTrue( VARIANT_TRUE, VT_BOOL );
	COleVariant vsFalse( VARIANT_FALSE, VT_BOOL );
		
	COleVariant vsNewBlankDocument( (LONG) 0 );

	MSWord::Documents objDocuments;

	m_objApplication.CreateDispatch(
			"Word.Application"
		);
		
	objDocuments = m_objApplication.GetDocuments();

	m_objDocument = objDocuments.Add(
			vsTemplate,
			vsFalse,
			vsNewBlankDocument,
			vsTrue
		);

	return;
}

/*****************************************************************************

 *****************************************************************************/

void cWRD_Document::close( void )
{
	COleVariant vsSaveChanges( (SHORT) -1 );
	COleVariant vsWordDocument( (SHORT) 0 );
	COleVariant vsFalse( VARIANT_FALSE, VT_BOOL );

	m_objApplication.Quit(
			(LPVARIANT) vsSaveChanges,
			(LPVARIANT) vsWordDocument,
			(LPVARIANT) vsFalse
		);

	return;
}

/*****************************************************************************


 *****************************************************************************/

cWRD_Table *cWRD_Document::getTable(
		LPCSTR p_sTableId
	)
{
	cWRD_Table *pTable;

	cWRD_mapTable::iterator itTemp;

	// 1. Look up table object.
	itTemp = m_mapTable.find(
			std::string(p_sTableId)
		);
		
	if (itTemp != m_mapTable.end())
	{
		pTable = itTemp->second;

		if (pTable != NULL)
		{
			return pTable;
		}
	}

	return NULL;
}

/*****************************************************************************


 *****************************************************************************/

void cWRD_Document::preprocess( void )
{
	_processDocument(
			m_objDocument,
			(pftExecuteMacro) &cWRD_Document::_execute_preprocess_macro_command
		);

	return;
}

/*****************************************************************************


 *****************************************************************************/

void cWRD_Document::render( void )
{
	_processDocument(
			m_objDocument,
			(pftExecuteMacro) &cWRD_Document::_execute_render_macro_command
		);

	return;
}

/*****************************************************************************



 *****************************************************************************/

void cWRD_Document::set( 
		char *p_sVarName,
		char *p_sValue
	)
{
	TRACE2("%s // %s \n", p_sVarName, p_sValue);

	m_mapSubstitute.insert(
			string_map::value_type(
					p_sVarName,
					p_sValue
				)
		);

	return;
}

/*****************************************************************************


 *****************************************************************************/

void cWRD_Document::setTableCell(
		char *p_sTableId,
		long p_lRowIndex,
		char *p_sColumnId,
		char *p_sValue
	)
{
	cWRD_Table *pTable;

	cWRD_mapTable::iterator itTemp;

	// 1. Look up table object.
	itTemp = m_mapTable.find(
			std::string(p_sTableId)
		);
		
	if (itTemp != m_mapTable.end())
	{
		pTable = itTemp->second;

		if (pTable != NULL)
		{
			pTable->set(
					p_lRowIndex,
					std::string(p_sColumnId),
					std::string(p_sValue)
				);
		}
	}
	
	return;
}

/*****************************************************************************


 *****************************************************************************/

LONG cWRD_Document::save( void )
{
	LONG lResult;

	try
	{
		m_objDocument.Save();

		lResult = S_OK;
	}
	catch(COleDispatchException *e)
	{
		lResult = e->m_scError;
	}
	catch(COleException *e)
	{
		lResult = COleException::Process(
				e
			);
	}
	catch(...)
	{
		lResult = E_FAIL;
	}

	return lResult;
}

/*****************************************************************************



 *****************************************************************************/

LONG cWRD_Document::saveas( 
		char *p_sPathName	   
	)
{
	LONG lResult;

	COleVariant vsPathName( p_sPathName );
	COleVariant vsNullString("");
	COleVariant vsFileFormat( (SHORT) 0 );

	COleVariant vsTrue( VARIANT_TRUE, VT_BOOL );
	COleVariant vsFalse( VARIANT_FALSE, VT_BOOL );

	m_objDocument.SaveAs(
			vsPathName,
			vsFileFormat,
			vsFalse,
			vsNullString,
			vsFalse,
			vsNullString,
			vsFalse,
			vsFalse,
			vsTrue,
			vsFalse,
			vsFalse
		);

	lResult = S_OK;

	return lResult;
}

/******************************************************************************

 ******************************************************************************/

LPDISPATCH cWRD_Document::_locate_word_idispatch_object_pointer(
		BSTR p_bstrType
	)
{
	BOOL bFound = false;

	UINT uiCount;

	BSTR bstrName;
	BSTR bstrDocString;
	DWORD dwordHelpContext;
	BSTR bstrHelpFile;

	LPDISPATCH lpDispatch_result;
	LPDISPATCH lpDispatch_temp;

	dispatch_stack stackDispatch;

	ITypeInfo *pTypeInfo;
	ITypeLib *pTypeLib;

	//	1.	Locate Table dispatch object
	while (!m_stackDispatch.empty() && !bFound)
	{
		lpDispatch_temp = m_stackDispatch.top();

		m_stackDispatch.pop();

		stackDispatch.push(
				lpDispatch_temp
			);

		lpDispatch_temp->GetTypeInfoCount(
				&uiCount
			);

		if ( uiCount > 0 )
		{
			lpDispatch_temp->GetTypeInfo(
					0,
					0,
					&pTypeInfo
				);

			pTypeInfo->GetContainingTypeLib(
					&pTypeLib,
					&uiCount
				);

			pTypeLib->GetDocumentation(
					uiCount,
					&bstrName,
					&bstrDocString,
					&dwordHelpContext,
					&bstrHelpFile
				);

			if (wcscmp(bstrName, p_bstrType) == 0 ) {				
				bFound = true;
				lpDispatch_result = lpDispatch_temp;
			}

			SysFreeString( 
					bstrName
				);

			SysFreeString(
					bstrDocString
				);

			SysFreeString(
					bstrHelpFile
				);

			pTypeLib->Release();

			pTypeInfo->Release();
		}
	}


	// put all the IDispatch interface pointers back on the stack
	while (!stackDispatch.empty())
	{
		lpDispatch_temp = stackDispatch.top();

		stackDispatch.pop();

		m_stackDispatch.push(
				lpDispatch_temp
			);
	}

	if (!bFound) {
		lpDispatch_result = NULL;
	}

	return lpDispatch_result;
}

/*****************************************************************************


 *****************************************************************************/

LPDISPATCH cWRD_Document::_locate_word_cell_idispatch_object_pointer(
		void
	)
{
	return _locate_word_idispatch_object_pointer(
			OLESTR("Cell")
		);
}


/*****************************************************************************


 *****************************************************************************/

LPDISPATCH cWRD_Document::_locate_word_row_idispatch_object_pointer( 
		void
	)
{
	return _locate_word_idispatch_object_pointer(
			OLESTR("Row")
		);
}

/*****************************************************************************


 *****************************************************************************/

cWRD_Object *cWRD_Document::_locate_stack_object(
		cWRD_Object::etType p_eType
	)
{
	BOOL bFound = FALSE;

	cWRD_Object *pObject_result;
	cWRD_Object *pObject_temp;

	cWRD_stackObject stackObject;

	while (!m_stackObject.empty() && !bFound)
	{
		pObject_temp = m_stackObject.top();

		m_stackObject.pop();

		stackObject.push(
				pObject_temp
			);

		if (pObject_temp->type() == p_eType)
		{
			bFound = true;
			pObject_result = pObject_temp;
		}
	}

	while (!stackObject.empty())
	{
		pObject_temp = stackObject.top();

		stackObject.pop();

		m_stackObject.push(
				pObject_temp
			);
	}

	if (!bFound)
	{
		pObject_result = NULL;
	}

	return pObject_result;
}

/*****************************************************************************


 *****************************************************************************/

LONG cWRD_Document::_determine_table_row_index_pair(
		LONG &p_rlRow_1,
		LONG &p_rlRow_2
	)
{
	LPDISPATCH lpDispatch;

	cWRD_Object *pTemp;

	cWRD_Row *pdRow_start;
	cWRD_Row *pdRow_end;

	if (!m_stackObject.empty())
	{
		pTemp = m_stackObject.top();

		m_stackObject.pop();

		if (pTemp->type() == cWRD_Object::etType::WORD_ROW)
		{
			pdRow_start = (cWRD_Row *) pTemp;

			lpDispatch = _locate_word_row_idispatch_object_pointer();

			if (lpDispatch != NULL)
			{
				pdRow_end = new cWRD_Row(
						lpDispatch
					);

				p_rlRow_1 = pdRow_start->getRowIndex();
				p_rlRow_2 = pdRow_end->getRowIndex();

				delete[] pdRow_end;
			}

			delete[] pdRow_start;
		}
	}

	return 0x00;
}

/*****************************************************************************


 *****************************************************************************/

LONG cWRD_Document::_execute_render_macro_command(
		etMode p_eMode,
		std::string p_sMacro,
		string_vector p_svParam,
		std::string &p_rsResult
	)
{	
	std::string sTemp;
	std::string sTemp2;

	string_map::iterator itTemp;

	if (p_sMacro == "substitute")
	{	
		sTemp = p_svParam[0];
		
		itTemp = m_mapSubstitute.find(
				sTemp
			);

		if (itTemp != m_mapSubstitute.end())
		{
			p_rsResult = (itTemp->second);

			return 0x01;
		}
	}

	p_rsResult = "";
	
	return 0x01;
}


/*****************************************************************************


 *****************************************************************************/

LONG cWRD_Document::_execute_preprocess_macro_command_table(
		std::string p_sMacro,
		string_vector p_svParam,
		std::string &p_rsResult
	)
{
	LPDISPATCH lpDispatch;

	cWRD_Table *pTable;

	lpDispatch = _locate_word_idispatch_object_pointer(
			OLESTR("Table")
		);

	if (lpDispatch != NULL)
	{
		pTable = new cWRD_Table(
				p_svParam[0], // name
				lpDispatch
			);

		m_mapTable.insert(
				cWRD_mapTable::value_type(
						p_svParam[0],
						pTable
					)
			);

		m_stackObject.push(
				pTable
			);
	}

	return 0x00;
}

/*****************************************************************************


 *****************************************************************************/

LONG cWRD_Document::_execute_preprocess_macro_command_table_header_begin(
		std::string p_sMacro,
		string_vector p_svParam,
		std::string &p_rsResult
	)
{
	LPDISPATCH pdRow;

	cWRD_Row *p_pRow;

	pdRow = _locate_word_row_idispatch_object_pointer();

	if (pdRow != NULL)
	{
		p_pRow = new cWRD_Row(
				pdRow
			);

		p_pRow->getRowIndex();

		m_stackObject.push(
				p_pRow
			);

		p_pRow = (cWRD_Row *) m_stackObject.top();

		p_pRow->getRowIndex();
	}

	return 0x00;
}

/*****************************************************************************


 *****************************************************************************/

LONG cWRD_Document::_execute_preprocess_macro_command_table_header_end(
		std::string p_sMacro,
		string_vector p_svParam,
		std::string &p_rsResult
	)
{
	LONG lRow_start;
	LONG lRow_end;

	cWRD_Object *pTemp;

	cWRD_Table *pdTable;

	if (!m_stackObject.empty())
	{
		_determine_table_row_index_pair(
				lRow_start,
				lRow_end
			);

		pTemp = m_stackObject.top();

		if (pTemp->type() == cWRD_Object::etType::WORD_TABLE)
		{
			pdTable = (cWRD_Table *) pTemp;

			pdTable->defineHeader(
					lRow_start,
					lRow_end
				);
		}
	}

	return 0x00;
}

/*****************************************************************************


 *****************************************************************************/

LONG cWRD_Document::_execute_preprocess_macro_command_table_body_begin(
		std::string p_sMacro,
		string_vector p_svParam,
		std::string &p_rsResult
	)
{
	LPDISPATCH pdRow;

	pdRow = _locate_word_row_idispatch_object_pointer();

	if (pdRow != NULL)
	{
		m_stackObject.push(
				new cWRD_Row(
						pdRow
					)
			);
	}

	return 0x00;
}

/*****************************************************************************


 *****************************************************************************/

LONG cWRD_Document::_execute_preprocess_macro_command_table_body_end(
		std::string p_sMacro,
		string_vector p_svParam,
		std::string &p_rsResult
	)
{
	LONG lRow_start;
	LONG lRow_end;

	cWRD_Object *pTemp;

	cWRD_Table *pdTable;

	if (!m_stackObject.empty())
	{
		_determine_table_row_index_pair(
				lRow_start,
				lRow_end
			);

		pTemp = m_stackObject.top();

		if (pTemp->type() == cWRD_Object::etType::WORD_TABLE)
		{
			pdTable = (cWRD_Table *) pTemp;

			pdTable->defineBody(
					lRow_start,
					lRow_end
				);
		}
	}

	return 0x00;
}

/*****************************************************************************


 *****************************************************************************/

LONG cWRD_Document::_execute_preprocess_macro_command_table_footer_begin(
		std::string p_sMacro,
		string_vector p_svParam,
		std::string &p_rsResult
	)
{
	LPDISPATCH pdRow;

	pdRow = _locate_word_row_idispatch_object_pointer();

	if (pdRow != NULL)
	{
		m_stackObject.push(
				new cWRD_Row(
						pdRow
					)
			);
	}

	return 0x00;
}

/*****************************************************************************


 *****************************************************************************/

LONG cWRD_Document::_execute_preprocess_macro_command_table_footer_end(
		std::string p_sMacro,
		string_vector p_svParam,
		std::string &p_rsResult
	)
{
	LONG lRow_start;
	LONG lRow_end;

	cWRD_Object *pTemp;

	cWRD_Table *pdTable;

	if (!m_stackObject.empty())
	{
		_determine_table_row_index_pair(
				lRow_start,
				lRow_end
			);

		pTemp = m_stackObject.top();

		if (pTemp->type() == cWRD_Object::etType::WORD_TABLE)
		{
			pdTable = (cWRD_Table *) pTemp;

			pdTable->defineFooter(
					lRow_start,
					lRow_end
				);
		}
	}

	return 0x00;
}

/*****************************************************************************


 *****************************************************************************/

LONG cWRD_Document::_execute_preprocess_macro_command_column_id(
		std::string p_sMacro,
		string_vector p_svParam,
		std::string &p_rsResult
	)
{
	LONG lColumn;

	LPDISPATCH pdCell;

	cWRD_Cell *pCell;
	cWRD_Table *pTable;

	pdCell = _locate_word_cell_idispatch_object_pointer();

	if (pdCell != NULL)
	{
		pCell = new cWRD_Cell(
				pdCell
			);

		lColumn = pCell->getColumnIndex();

		pTable = (cWRD_Table *) _locate_stack_object(
				cWRD_Object::etType::WORD_TABLE
			);

		if (pTable != NULL)
		{
			pTable->defineColumnId(
					p_svParam[0],
					lColumn
				);
		}

		delete[] pCell;
	}

	return 0x00;
}

/*****************************************************************************


 *****************************************************************************/

LONG cWRD_Document::_execute_preprocess_macro_command(
		etMode p_eMode,
		std::string p_sMacro,
		string_vector p_svParam,
		std::string &p_rsResult
	)
{
	switch (p_eMode)
	{
	case modeTable:

		if (p_sMacro == "table-id")
		{
			// execute table create, using object on stack
			return _execute_preprocess_macro_command_table(
					p_sMacro,
					p_svParam,
					p_rsResult
				);
		}

		if (p_sMacro == "table-header-begin")
		{
			return _execute_preprocess_macro_command_table_header_begin(
					p_sMacro,
					p_svParam,
					p_rsResult
				);
		}

		if (p_sMacro == "table-header-end")
		{
			return _execute_preprocess_macro_command_table_header_end(
					p_sMacro,
					p_svParam,
					p_rsResult
				);
		}

		if (p_sMacro == "table-body-begin")
		{
			return _execute_preprocess_macro_command_table_body_begin(
					p_sMacro,
					p_svParam,
					p_rsResult
				);
		}

		if (p_sMacro == "table-body-end")
		{
			return _execute_preprocess_macro_command_table_body_end(
					p_sMacro,
					p_svParam,
					p_rsResult
				);
		}

		if (p_sMacro == "table-footer-begin")
		{
			return _execute_preprocess_macro_command_table_footer_begin(
					p_sMacro,
					p_svParam,
					p_rsResult
				);
		}

		if (p_sMacro == "table-footer-end")
		{
			return _execute_preprocess_macro_command_table_footer_end(
					p_sMacro,
					p_svParam,
					p_rsResult
				);
		}

		if (p_sMacro == "column-id")
		{
			return _execute_preprocess_macro_command_column_id(
					p_sMacro,
					p_svParam,
					p_rsResult
				);
		}

		break;
	}

	return 0x00;
}

/*****************************************************************************


 *****************************************************************************/

void cWRD_Document::_processCellSet(
		MSWord::Cells &p_robjCells,
		pftExecuteMacro p_pftExecuteMacro
	)
{
	BOOL blResult;

	LONG lResult;

	LONG lPos;

	LONG lPos_start;
	LONG lPos_end;

	LONG lRange_start;

	LONG lCell;
	LONG lCell_max;

	MSWord::Cell objCell;
	MSWord::Range objRange;
	MSWord::Selection objSelection;

	std::string sMacro;
	string_vector vsParam;

	m_stackDispatch.push(
			p_robjCells.m_lpDispatch
		);

	lCell_max = p_robjCells.GetCount();

	for (lCell = 0x01; lCell <= lCell_max; lCell++)
	{
		objCell = p_robjCells.Item(lCell);

		m_stackDispatch.push(
				objCell.m_lpDispatch
			);

		objRange = objCell.GetRange();

		std::string sText = objRange.GetText();
		std::string sTemp;

		lPos = 0;

		lRange_start = objRange.GetStart();

		blResult = _seekMacro(
				sText,
				lPos,
				sMacro,
				vsParam,
				lPos_start,
				lPos_end
			);
		
		while (blResult)
		{
			objRange.SetRange( 
					lRange_start + lPos_start, 
					lRange_start + lPos_end + 1
				);

			lResult = (this->*p_pftExecuteMacro)(
					modeTable,
					sMacro,
					vsParam,
					sTemp
				);

			TRACE2("'%s', '%s' \n", sText.c_str(), sTemp.c_str() );

			if (lResult)
			{
				COleVariant vsCollapseStart( (SHORT) 1 );
				COleVariant vsCharacter( (SHORT) 1 );
				COleVariant vsDel( (LONG) 0 );
			
				objRange.Delete(
						vsCharacter,
						vsDel
					);
				
				objRange.SetText(
						sTemp.c_str()
					);
			}

			objRange = objCell.GetRange();
			sText = objRange.GetText();
			lRange_start = objRange.GetStart();

			blResult = _seekMacro(
					sText,
					lPos,
					sMacro,
					vsParam,
					lPos_start,
					lPos_end
				);
		}

		m_stackDispatch.pop();
	}

	m_stackDispatch.pop();

	return;
}

/*****************************************************************************


 *****************************************************************************/

void cWRD_Document::_processDocument(
		 MSWord::_Document &p_robjDocument,
		 pftExecuteMacro p_pftExecuteMacro
	)
{
	MSWord::Frames objFrames;
	MSWord::Shapes objShapes;

	m_stackDispatch.push(
			p_robjDocument.m_lpDispatch
		);

	objFrames = p_robjDocument.GetFrames();
	objShapes = p_robjDocument.GetShapes();

	_processFrameSet(
			objFrames,
			p_pftExecuteMacro
		);

	_processShapeSet(
			objShapes,
			p_pftExecuteMacro
		);

	m_stackDispatch.pop();

	return;
}

/*****************************************************************************


 *****************************************************************************/

void cWRD_Document::_processFrame(
		MSWord::Frame &p_robjFrame,
		pftExecuteMacro p_pftExecuteMacro
	)
{
	MSWord::Range objRange;
	MSWord::Tables objTables;

	m_stackDispatch.push(
			p_robjFrame.m_lpDispatch
		);

	objRange = p_robjFrame.GetRange();

	objTables = objRange.GetTables();

	if (objTables.GetCount() == 0)
	{
		_processRange(
				objRange,
				p_pftExecuteMacro
			);
	}
	else
	{
		_processTableSet(
				objTables,
				p_pftExecuteMacro
			);
	}

	m_stackDispatch.pop();

	return;
}

/*****************************************************************************


 *****************************************************************************/

void cWRD_Document::_processFrameSet(
		 MSWord::Frames &p_robjFrames,
		 pftExecuteMacro p_pftExecuteMacro
	)
{
	LONG i;
	LONG i_max;

	m_stackDispatch.push(
			p_robjFrames.m_lpDispatch
		);			

	i_max = p_robjFrames.GetCount();

	MSWord::Frame objFrame;

	for (i = 1; i <= i_max; i++)
	{
		objFrame = p_robjFrames.Item(i);

		_processFrame(
				objFrame,
				p_pftExecuteMacro
			);
	}

	m_stackDispatch.pop();

	return;
}

/******************************************************************************


 ******************************************************************************/

void cWRD_Document::_processParagraph(
		MSWord::Paragraph &p_robjParagraph,
		pftExecuteMacro p_pftExecuteMacro
	)
{
	BOOL blResult;

	LONG lResult;

	LONG lPos;
	LONG lPos_start;
	LONG lPos_end;

	LONG lRange_start;

	std::string sText;
	std::string sTemp;

	std::string sMacro;
	string_vector vsParam;

	MSWord::Range objRange;

	m_stackDispatch.push(
			p_robjParagraph.m_lpDispatch
		);

	objRange = p_robjParagraph.GetRange();

	sText = objRange.GetText();

	lPos = 0;

	lRange_start = objRange.GetStart();

	std::string nnn;

	blResult = _seekMacro(
			sText,
			lPos,
			sMacro,
			vsParam,
			lPos_start,
			lPos_end
		);
		
	while (blResult)
	{
		objRange.SetRange( 
				lRange_start + lPos_start, 
				lRange_start + lPos_end + 1
			);

		nnn = objRange.GetText();

		TRACE1(" xxxx -> '%s'\n", nnn.c_str());

		lResult = (this->*p_pftExecuteMacro)(
				modeParagraph,
				sMacro,
				vsParam,
				sTemp
			);

		TRACE2("'%s', '%s' \n", sText.c_str(), sTemp.c_str() );
		
		if (lResult)
		{		
			COleVariant vsCollapseStart( (SHORT) 1 );
			COleVariant vsCharacter( (SHORT) 1 );
			COleVariant vsDel( (LONG) 0 );
			
			objRange.Delete(
					vsCharacter,
					vsDel
				);
			
			objRange.SetText(
					sTemp.c_str()
				);
		}

		objRange = p_robjParagraph.GetRange();
		sText = objRange.GetText();
		lRange_start = objRange.GetStart();

		blResult = _seekMacro(
				sText,
				lPos,
				sMacro,
				vsParam,
				lPos_start,
				lPos_end
			);
	}

	/*sTemp = _parse(
			sText,
			modeParagraph,
			p_pftExecuteMacro
		);

	TRACE2("'%s', '%s' \n", sText.c_str(), sTemp.c_str() );

	if (sText != sTemp)
	{
		objRange.SetText(
				sTemp.c_str()
			);
	}*/

	m_stackDispatch.pop();

	return;
}

/******************************************************************************


 ******************************************************************************/

void cWRD_Document::_processParagraphSet(
		MSWord::Paragraphs &p_robjParagraphs,
		pftExecuteMacro p_pftExecuteMacro
	)
{
	LONG lParagraph;
	LONG lParagraph_max;

	MSWord::Paragraph objParagraph;

	m_stackDispatch.push(
			p_robjParagraphs.m_lpDispatch
		);

	lParagraph_max = p_robjParagraphs.GetCount();

	for (lParagraph = 0x01; lParagraph <= lParagraph_max; lParagraph++)
	{
		objParagraph = p_robjParagraphs.Item(lParagraph);

		_processParagraph(
				objParagraph,
				p_pftExecuteMacro
			);
	}

	m_stackDispatch.pop();

	return;
}

/*****************************************************************************


 *****************************************************************************/

void cWRD_Document::_processRange(
		MSWord::Range &p_robjRange,
		pftExecuteMacro p_pftExecuteMacro
	)
{
	MSWord::Tables objTables;
	MSWord::Paragraphs objParagraphs;

	m_stackDispatch.push(
			p_robjRange.m_lpDispatch
		);

	objTables = p_robjRange.GetTables();

	if (objTables != NULL) 
	{
		_processTableSet(
				objTables,
				p_pftExecuteMacro
			);
	}

	objParagraphs = p_robjRange.GetParagraphs();

	if (objParagraphs != NULL)
	{
		_processParagraphSet(
				objParagraphs,
				p_pftExecuteMacro
			);
	}

	m_stackDispatch.pop();

	return;
}

/*****************************************************************************


 *****************************************************************************/

void cWRD_Document::_processRowSet(
		MSWord::Rows &p_robjRows,
		pftExecuteMacro p_pftExecuteMacro
	)
{
	LONG lRow;
	LONG lRow_max;

	MSWord::Row objRow;
	MSWord::Cells objCells;

	m_stackDispatch.push(
			p_robjRows.m_lpDispatch
		);

	lRow_max = p_robjRows.GetCount();

	for (lRow = 0x01; lRow <= lRow_max; lRow++)
	{
		objRow = p_robjRows.Item(lRow);

		m_stackDispatch.push(
				objRow.m_lpDispatch
			);

		objCells = objRow.GetCells();

		_processCellSet(
				objCells,
				p_pftExecuteMacro
			);

		m_stackDispatch.pop();
	}

	m_stackDispatch.pop();

	return;
}

/******************************************************************************


 ******************************************************************************/

void cWRD_Document::_processShape(
		MSWord::Shape &p_robjShape,
		pftExecuteMacro p_pftExecuteMacro
	)
{
	MSWord::TextFrame objTextFrame;
	MSWord::Range objRange;

	m_stackDispatch.push(
			p_robjShape.m_lpDispatch
		);

	objTextFrame = p_robjShape.GetTextFrame();
	
	if (objTextFrame != NULL) 
	{
		if (objTextFrame.GetHasText()) 
		{
			objRange = objTextFrame.GetTextRange();		

			if (objRange != NULL) 
			{
				_processRange(
						objRange,
						p_pftExecuteMacro
					);
			}
		}
	}

	m_stackDispatch.pop();

	return;
}

/******************************************************************************


 ******************************************************************************/

void cWRD_Document::_processShapeSet(
		MSWord::Shapes &p_robjShapes,
		pftExecuteMacro p_pftExecuteMacro
	)
{
	LONG lShape;
	LONG lShape_max;

	VARIANT vtTemp;

	MSWord::Shape objShape;

	m_stackDispatch.push(
			p_robjShapes.m_lpDispatch
		);

	lShape_max = p_robjShapes.GetCount();

	for (lShape = 0x01; lShape <= lShape_max; lShape++)
	{
		vtTemp.lVal = lShape;
		vtTemp.vt = VT_I4;

		objShape = p_robjShapes.Item(&vtTemp);

		_processShape(
				objShape,
				p_pftExecuteMacro
			);
	}

	m_stackDispatch.pop();

	return;
}

/*****************************************************************************


 *****************************************************************************/

void cWRD_Document::_processTable(
		MSWord::Table &p_robjTable,
		pftExecuteMacro p_pftExecuteMacro
	)
{
	return;
}

/*****************************************************************************


 *****************************************************************************/

void cWRD_Document::_processTableSet(
		MSWord::Tables &p_robjTables,
		pftExecuteMacro p_pftExecuteMacro
	)
{
	LONG lTable;
	LONG lTable_max;

	MSWord::Table objTable;
	MSWord::Rows objRows;

	m_stackDispatch.push(
			p_robjTables.m_lpDispatch
		);

	lTable_max = p_robjTables.GetCount();

	for (lTable = 0x01; lTable <= lTable_max; lTable++)
	{
		objTable = p_robjTables.Item(lTable);

		m_stackDispatch.push(
				objTable.m_lpDispatch
			);

		objRows = objTable.GetRows();
		
		_processRowSet(
				objRows,
				p_pftExecuteMacro
			);

		m_stackDispatch.pop();
	}

	m_stackDispatch.pop();

	return;
}

/*****************************************************************************


 *****************************************************************************/

BOOL cWRD_Document::_test_word_idispatch_object_pointer(
		LPDISPATCH p_lpDispatch,
		BSTR p_bstrType
	)
{
		BOOL bFound = false;

	UINT uiCount;

	BSTR bstrName;
	BSTR bstrDocString;
	DWORD dwordHelpContext;
	BSTR bstrHelpFile;

	dispatch_stack stackDispatch;

	ITypeInfo *pTypeInfo;
	ITypeLib *pTypeLib;

	p_lpDispatch->GetTypeInfoCount(
			&uiCount
		);

	if ( uiCount > 0 )
	{
		p_lpDispatch->GetTypeInfo(
				0,
				0,
				&pTypeInfo
			);

		pTypeInfo->GetContainingTypeLib(
				&pTypeLib,
				&uiCount
			);

		pTypeLib->GetDocumentation(
				uiCount,
				&bstrName,
				&bstrDocString,
				&dwordHelpContext,
				&bstrHelpFile
			);

		if (wcscmp(bstrName, p_bstrType) == 0 ) {				
			bFound = true;
		}

		SysFreeString( 
				bstrName
			);

		SysFreeString(
				bstrDocString
			);

		SysFreeString(
				bstrHelpFile
			);

		pTypeLib->Release();

		pTypeInfo->Release();
	}
	
	return bFound;
}

/*****************************************************************************

 End of Module

 *****************************************************************************/