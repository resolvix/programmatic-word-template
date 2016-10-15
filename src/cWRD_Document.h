// cWRD_Document.h: interface for the cWRD_Document class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CWRD_DOCUMENT_H__7FBC0079_C0E9_4471_B02A_A2537C12D28B__INCLUDED_)
#define AFX_CWRD_DOCUMENT_H__7FBC0079_C0E9_4471_B02A_A2537C12D28B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class cWRD_setTable;

class cWRD_Document : public cGEN_MacroParser
{
public:

	cWRD_Document();

	virtual ~cWRD_Document();

	void close( void );

	void create(
			char *p_sTemplate			 			 
		);

	cWRD_Table *getTable(
			LPCSTR p_sTableId
		);

	void preprocess( void );

	void render( void );

	void set( 	
			char *p_sVarName,
			char *p_sValue
		);

	void setTableCell(
			char *p_sTableId,
			long p_lRowIndex,
			char *p_sColumnId,
			char *p_sValue
		);

	LONG save( void );

	LONG saveas( 
			char *p_sPathName
		);

protected:

	typedef LONG (cWRD_Document::*pftExecuteMacro)(
			etMode p_eMode,
			std::string p_sMacro,
			string_vector p_svParam,
			std::string &p_rsResult
		);

	typedef std::stack< 
			LPDISPATCH, 
			std::deque< 
					LPDISPATCH,
					std::allocator< LPDISPATCH >
				>
		> dispatch_stack;

	MSWord::_Application m_objApplication;
	MSWord::_Document m_objDocument;

	cWRD_stackObject m_stackObject;

	cWRD_mapTable m_mapTable;

	string_map m_mapSubstitute;
	
	dispatch_stack m_stackDispatch;

	LPDISPATCH _locate_word_idispatch_object_pointer(
			BSTR p_bstrType
		);

	LPDISPATCH _locate_word_cell_idispatch_object_pointer(
			void
		);

	LPDISPATCH _locate_word_row_idispatch_object_pointer( 
			void
		);

	cWRD_Object *_locate_stack_object(
			cWRD_Object::etType p_eType
		);

	LONG _determine_table_row_index_pair(
			LONG &p_rlRow_1,
			LONG &p_rlRow_2
		);

	LONG _execute_preprocess_macro_command(
			etMode p_eMode,
			std::string p_sMacro,
			string_vector p_svParam,
			std::string &p_rsResult
		);

	LONG _execute_render_macro_command(
			etMode p_eMode,
			std::string p_sMacro,
			string_vector p_svParam,
			std::string &p_rsResult
		);

	LONG _execute_preprocess_macro_command_table(
			std::string p_sMacro,
			string_vector p_svParam,
			std::string &p_rsResult
		);

	LONG _execute_preprocess_macro_command_table_header_begin(
			std::string p_sMacro,
			string_vector p_svParam,
			std::string &p_rsResult
		);

	LONG _execute_preprocess_macro_command_table_header_end(
			std::string p_sMacro,
			string_vector p_svParam,
			std::string &p_rsResult
		);

	LONG _execute_preprocess_macro_command_table_body_begin(
			std::string p_sMacro,
			string_vector p_svParam,
			std::string &p_rsResult
		);

	LONG _execute_preprocess_macro_command_table_body_end(
			std::string p_sMacro,
			string_vector p_svParam,
			std::string &p_rsResult
		);

	LONG _execute_preprocess_macro_command_table_footer_begin(
			std::string p_sMacro,
			string_vector p_svParam,
			std::string &p_rsResult
		);

	LONG _execute_preprocess_macro_command_table_footer_end(
			std::string p_sMacro,
			string_vector p_svParam,
			std::string &p_rsResult
		);

	LONG _execute_preprocess_macro_command_column_id(
			std::string p_sMacro,
			string_vector p_svParam,
			std::string &p_rsResult
		);

	void _processCellSet(
			MSWord::Cells &p_robjCells,
			pftExecuteMacro p_pftExecuteMacro
		);

	void _processDocument(
			MSWord::_Document &p_robjDocument,
			pftExecuteMacro p_pftExecuteMacro
		);

	void _processFrame(
			MSWord::Frame &p_robjFrame,
			pftExecuteMacro p_pftExecuteMacro
		);

	void _processFrameSet(
			MSWord::Frames &p_robjFrames,
			pftExecuteMacro p_pftExecuteMacro
		);

	void _processParagraph(
			MSWord::Paragraph &p_robjParagraph,
			pftExecuteMacro p_pftExecuteMacro
		);

	void _processParagraphSet(
			MSWord::Paragraphs &p_robjParagraphs,
			pftExecuteMacro p_pftExecuteMacro
		);

	void _processRange(
			MSWord::Range &p_robjRange,
			pftExecuteMacro p_pftExecuteMacro
		);
	
	void _processRowSet(
			MSWord::Rows &p_robjRows,
			pftExecuteMacro p_pftExecuteMacro
		);

	void _processShape(
			MSWord::Shape &p_robjShape,
			pftExecuteMacro p_pftExecuteMacro
		);

	void _processShapeSet(
			MSWord::Shapes &p_robjShapes,
			pftExecuteMacro p_pftExecuteMacro
		);

	void _processTable(
			MSWord::Table &p_robjTable,
			pftExecuteMacro p_pftExecuteMacro
		);

	void _processTableSet(
			MSWord::Tables &p_robjTables,
			pftExecuteMacro p_pftExecuteMacro
		);

	BOOL _test_word_idispatch_object_pointer(
			LPDISPATCH p_lpDispatch,
			BSTR p_bstrType
		);

};

#endif // !defined(AFX_CWRD_DOCUMENT_H__7FBC0079_C0E9_4471_B02A_A2537C12D28B__INCLUDED_)
