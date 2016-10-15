/*****************************************************************************

 *****************************************************************************/

#include "stdafx.h"

#include "cGEN_ErrorHandler.h"

/*****************************************************************************

 *****************************************************************************/

HRESULT cGEN_ErrorHandler::ComException(
		CException *p_pException,
		REFGUID p_GUID
	)
{
	HRESULT hrTemp;
	HRESULT hResult;
		
	ICreateErrorInfo *pCreateErrorInfo;
	IErrorInfo *pErrorInfo;

	hrTemp = ::CreateErrorInfo(
			&pCreateErrorInfo
		);

	if (SUCCEEDED(hrTemp))
	{

		if (p_pException->IsKindOf( RUNTIME_CLASS(COleDispatchException) ))
		{
			COleDispatchException *pE = (COleDispatchException *) p_pException;

			pCreateErrorInfo->SetGUID(
					p_GUID
				);

			pCreateErrorInfo->SetDescription(
					pE->m_strDescription.AllocSysString()
				);

			pCreateErrorInfo->SetSource( 
					pE->m_strSource.AllocSysString()
				);

			pCreateErrorInfo->SetHelpContext(
					pE->m_dwHelpContext
				);

			pCreateErrorInfo->SetHelpFile(
					pE->m_strHelpFile.AllocSysString()
				);

			hResult = pE->m_scError;
		}
		else
		{
			// do nothing for the time being

		}
		
	}

	hrTemp = pCreateErrorInfo->QueryInterface(
			IID_IErrorInfo,
			(LPVOID *) &pErrorInfo
		);

	if (SUCCEEDED(hrTemp))
	{
		::SetErrorInfo(
				0, 
				pErrorInfo
			);

		pCreateErrorInfo->Release();
	}

	pErrorInfo->Release();

	p_pException->Delete();

	return hResult;
}

/*****************************************************************************

 *****************************************************************************/