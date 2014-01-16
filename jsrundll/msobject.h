#pragma once
#include "jsrundll_h.h"
#include "IDispatchImpl2.h"

class __declspec(uuid("{F8DB2C1A-AF61-438F-8D7A-0C76482CEFC6}")) CMs
	: public CComObjectRootEx<CComSingleThreadModel>
	, public CComCoClass<CMs, &__uuidof(CMs)>
	, public IDispatchImpl2<IObjectMs>
{
public:
	CMs(){}
	~CMs(){}
	BEGIN_COM_MAP(CMs)
		COM_INTERFACE_ENTRY(IObjectMs)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP()

	DECLARE_NO_REGISTRY()
	DECLARE_CLASSFACTORY()
	DECLARE_PROTECT_FINAL_CONSTRUCT()

	STDMETHODIMP XMLHttpRequest( VARIANT* ret )
	{
		CComPtr<IXMLHttpRequest> req;
	
		CLSID clsid;
		HRESULT hr = E_FAIL;
		hr = CLSIDFromProgID( L"MSXML2.XMLHTTP", &clsid );	
		if ( SUCCEEDED(hr) )
		{
			hr = req.CoCreateInstance( clsid,NULL );
			if ( SUCCEEDED(hr) )
			{
				ret->vt = VT_DISPATCH;
				ret->pdispVal = req;
				ret->pdispVal->AddRef();
			}
		}
		return hr;
	}
	STDMETHODIMP XMLDOMDocument( VARIANT* ret )
	{
		CLSID clsid;
		HRESULT hr = E_FAIL;
		if ( S_OK!=CLSIDFromProgID( L"MSXML2.DOMDocument.6.0", &clsid ))
			 hr =CLSIDFromProgID( L"MSXML2.DOMDocument", &clsid );

		if ( SUCCEEDED(hr))
		{
			CComPtr<IXMLDOMDocument> xml;
			hr = xml.CoCreateInstance( clsid,NULL );
			if ( SUCCEEDED(hr) )
			{
				CComPtr<IDispatch> xd;
				hr = xml.QueryInterface( &xd );

				if ( SUCCEEDED(hr) )
				{
					ret->vt = VT_DISPATCH;
					ret->pdispVal = xd;
					ret->pdispVal->AddRef();
				}
			}
		}
		return hr;
	}
	STDMETHODIMP Guid( VARIANT* ret )
	{
		GUID uuid; 			
		CoCreateGuid(&uuid);
				
		ATL::CString r;
		r.Format( L"%.8x-%.4x-%.4x-%.2x%.2x-%.2x%.2x%.2x%.2x%.2x%.2x", uuid.Data1, uuid.Data2, uuid.Data3, uuid.Data4[0],uuid.Data4[1],
					uuid.Data4[2],uuid.Data4[3],uuid.Data4[4],uuid.Data4[5],uuid.Data4[6],uuid.Data4[7] );		

		CComVariant vr = r;
		*ret = vr;

		return S_OK;
	}

	static LPCWSTR name(){ return L"ms"; }
};


