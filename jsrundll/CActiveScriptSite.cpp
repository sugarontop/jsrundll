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
	if (!pDebugDoc_) 
		return E_UNEXPECTED;

	ULONG ulStartPos=0; 

	if ( S_OK == pDebugDoc_->GetScriptBlockInfo(pdwSourceContext, 0, &ulStartPos, 0))
		return pDebugDoc_->CreateDebugDocumentContext(ulStartPos+uCharacterOffset, uNumChars, ppsc); 

	return E_FAIL;
} 

STDMETHODIMP CActiveScriptSite::GetApplication(/*[out]*/ IDebugApplication **ppda) 
{ 
	if (!ppda) 
		return E_POINTER; 

	*ppda = 0; 
	if (!pDebugApp_) 
		return E_UNEXPECTED; 

	*ppda = pDebugApp_; 
	pDebugApp_->AddRef();
	return S_OK; 
} 

STDMETHODIMP CActiveScriptSite::GetRootApplicationNode(/*[out]*/ IDebugApplicationNode** ppdanRoot) 
{ 
	if (!ppdanRoot) 
		return E_POINTER; 

	*ppdanRoot = 0; 
	return S_OK; 
}

STDMETHODIMP CActiveScriptSite::OnScriptErrorDebug(/*[in]*/ IActiveScriptErrorDebug* pErrorDebug,
											  /*[out]*/ BOOL *pfEnterDebugger,
											  /*[out]*/ BOOL *pfCallOnScriptErrorWhenContinuing) 
{ 
	
	*pfEnterDebugger = FALSE; 
	*pfCallOnScriptErrorWhenContinuing = TRUE;
	return S_OK; 
}

HRESULT CActiveScriptSite::CreateScriptDebugger()
{
	HRESULT hr = ::CoCreateInstance(CLSID_ProcessDebugManager, NULL, CLSCTX_INPROC_SERVER, 
		IID_IProcessDebugManager, (void**)&pDebugMgr_);
	if (SUCCEEDED(hr))
	{
		hr = pDebugMgr_->CreateApplication(&pDebugApp_);
	}
	if (SUCCEEDED(hr))
	{
		hr = pDebugApp_->SetName(L"jsrun");
	}
	if (SUCCEEDED(hr))
	{
		hr = pDebugMgr_->AddApplication(pDebugApp_, &AppCookie_);
	}

	if (FAILED(hr))
	{
		ReleaseScriptDebugger();
	}
	return hr;
}



HRESULT CActiveScriptSite::CreateDocumentForDebugger(BSTR scripts)
{
	if (!pDebugMgr_)
		return E_UNEXPECTED;
	
	ReleaseDebugDocument();
	
	HRESULT hr = pDebugMgr_->CreateDebugDocumentHelper(0, &pDebugDoc_);
	if (SUCCEEDED(hr))
	{
		hr = pDebugDoc_->Init(pDebugApp_, L"short codename", L"long codename", TEXT_DOC_ATTR_READONLY|TEXT_DOC_ATTR_TYPE_SCRIPT);
	}
	
	
	if (SUCCEEDED(hr))
	{
		hr = pDebugDoc_->Attach(0);
	}

	/*if (SUCCEEDED(hr))
	{
		dochost.Release();
		hr = DebugDocumentHost::CreateInstance( &dochost );
		
		hr = pDebugDoc_->SetDebugDocumentHost( dochost );

	}*/


	if (SUCCEEDED(hr))
	{
		hr = pDebugDoc_->AddUnicodeText(scripts);
	}
	if (SUCCEEDED(hr))
	{
		ULONG len = ::SysStringLen(scripts);
		hr = pDebugDoc_->DefineScriptBlock(0, len, pAS, FALSE, &pdwSourceContext);
	}
	
	/*if (SUCCEEDED(hr))
	{
		hr = pDebugDoc_->BringDocumentToTop();

		CComPtr<IDebugApplication> app;
		GetApplication( &app );

		hr = app->CauseBreak();
	}*/

	

	return hr;
}

void CActiveScriptSite::ReleaseDebugDocument()
{
	if (pDebugDoc_)
	{
		pDebugDoc_->Detach();
		pDebugDoc_->Release();
		pDebugDoc_ = 0;
	}
}

void CActiveScriptSite::ReleaseScriptDebugger()
{
	// Release script debugger interfaces.
	ReleaseDebugDocument();

	if (pDebugMgr_ && AppCookie_)
	{
		pDebugMgr_->RemoveApplication(AppCookie_);
		AppCookie_ = 0;
	}
	
	if ( pDebugApp_ )
	{
		pDebugApp_->Release();
		pDebugApp_ = NULL;
	}
	
	if ( pDebugMgr_ )
	{
		pDebugMgr_->Release();
		pDebugMgr_ = NULL;
	}
}

#endif // __Script_Debug__


