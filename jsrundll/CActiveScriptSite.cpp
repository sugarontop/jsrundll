#include "stdafx.h"
#include "CActiveScriptSite.h"


extern CComPtr<IActiveScript> pAS;

// For script debugging...
#ifdef __Script_Debug__
// --- IActiveScriptSiteDebug methods ---
STDMETHODIMP CActiveScriptSite::GetDocumentContextFromPosition(/*[in]*/ DWORD dwSourceContext,
														  /*[in]*/ ULONG uCharacterOffset,
														  /*[in]*/ ULONG uNumChars,
														  /*[out]*/ IDebugDocumentContext** ppsc) 
{ 	
	dwSourceContext;
	if (!ppsc) 
		return E_POINTER;

	*ppsc = 0; 
	if (!mpDebugDoc) 
		return E_UNEXPECTED;

	ULONG ulStartPos=0; 

	if ( S_OK == mpDebugDoc->GetScriptBlockInfo(pdwSourceContext, 0, &ulStartPos, 0))
		return mpDebugDoc->CreateDebugDocumentContext(ulStartPos+uCharacterOffset, uNumChars, ppsc); 

	return E_FAIL;
} 

STDMETHODIMP CActiveScriptSite::GetApplication(/*[out]*/ IDebugApplication **ppda) 
{ 
	if (!ppda) 
		return E_POINTER; 

	*ppda = 0; 
	if (!mpDebugApp) 
		return E_UNEXPECTED; 

	*ppda = mpDebugApp; 
	mpDebugApp->AddRef();
	return S_OK; 
} 

STDMETHODIMP CActiveScriptSite::GetRootApplicationNode(/*[out]*/ IDebugApplicationNode** ppdanRoot) 
{ 
	// If we have only one document, we can safely return NULL here (this is the root). 
	if (!ppdanRoot) 
		return E_POINTER; 

	*ppdanRoot = 0; 
	return S_OK; 
}

STDMETHODIMP CActiveScriptSite::OnScriptErrorDebug(/*[in]*/ IActiveScriptErrorDebug* pErrorDebug,
											  /*[out]*/ BOOL *pfEnterDebugger,
											  /*[out]*/ BOOL *pfCallOnScriptErrorWhenContinuing) 
{ 
	pErrorDebug;
	// Runtime errors get here before go into IActiveScriptSite::OnScriptError 
	// Query the IActiveScriptErrorDebug interface for more 
	// info on what kind of error occurred. 
	*pfEnterDebugger = FALSE; 
	*pfCallOnScriptErrorWhenContinuing = TRUE;
	return S_OK; 
}

HRESULT CActiveScriptSite::CreateScriptDebugger()
{
	// Create script debugger interfaces.
	HRESULT hr = ::CoCreateInstance(CLSID_ProcessDebugManager, NULL, CLSCTX_INPROC_SERVER, 
		IID_IProcessDebugManager, (void**)&mpDebugMgr);
	if (SUCCEEDED(hr))
	{
		hr = mpDebugMgr->CreateApplication(&mpDebugApp);
	}
	if (SUCCEEDED(hr))
	{
		hr = mpDebugApp->SetName(L"jsrun");
	}
	if (SUCCEEDED(hr))
	{
		hr = mpDebugMgr->AddApplication(mpDebugApp, &mAppCookie);
	}

	if (FAILED(hr))
	{
		ReleaseScriptDebugger();
	}
	return hr;
}



HRESULT CActiveScriptSite::CreateDocumentForDebugger(BSTR scripts)
{
	if (!mpDebugMgr)
		return E_UNEXPECTED;

	// Release previous used document.
	ReleaseDebugDocument();
	
	HRESULT hr = mpDebugMgr->CreateDebugDocumentHelper(0, &mpDebugDoc);
	if (SUCCEEDED(hr))
	{
		hr = mpDebugDoc->Init(mpDebugApp, L"short codename", L"long codename", TEXT_DOC_ATTR_READONLY|TEXT_DOC_ATTR_TYPE_SCRIPT);
	}
	
	
	if (SUCCEEDED(hr))
	{
		hr = mpDebugDoc->Attach(0);
	}

	/*if (SUCCEEDED(hr))
	{
		dochost.Release();
		hr = DebugDocumentHost::CreateInstance( &dochost );
		
		hr = mpDebugDoc->SetDebugDocumentHost( dochost );

	}*/


	if (SUCCEEDED(hr))
	{
		hr = mpDebugDoc->AddUnicodeText(scripts);
	}
	if (SUCCEEDED(hr))
	{
		ULONG len = ::SysStringLen(scripts);
		hr = mpDebugDoc->DefineScriptBlock(0, len, pAS, FALSE, &pdwSourceContext);
	}
	
	/*if (SUCCEEDED(hr))
	{
		hr = mpDebugDoc->BringDocumentToTop();

		CComPtr<IDebugApplication> app;
		GetApplication( &app );

		hr = app->CauseBreak();
	}*/

	

	return hr;
}

void CActiveScriptSite::ReleaseDebugDocument()
{
	if (mpDebugDoc)
	{
		mpDebugDoc->Detach();
		mpDebugDoc->Release();
		mpDebugDoc = 0;
	}
}

void CActiveScriptSite::ReleaseScriptDebugger()
{
	// Release script debugger interfaces.
	ReleaseDebugDocument();

	if (mpDebugMgr && mAppCookie)
	{
		mpDebugMgr->RemoveApplication(mAppCookie);
		mAppCookie = 0;
	}
	//SAFE_RELEASE(mpDebugApp);
	if ( mpDebugApp )
	{
		mpDebugApp->Release();
		mpDebugApp = NULL;
	}
	//SAFE_RELEASE(mpDebugMgr);

	if ( mpDebugMgr )
	{
		mpDebugMgr->Release();
		mpDebugMgr = NULL;
	}
}

#endif // __Script_Debug__


