#pragma once

#include "jsrundll_h.h"
#include "IDispatchImpl2.h"


class __declspec(uuid("{15F47057-DE9B-4638-8B72-CABF72E93658}")) CDebug
	: public CComObjectRootEx<CComSingleThreadModel>
	, public CComCoClass<CDebug, &__uuidof(CDebug)>
	, public IDispatchImpl2<IObjectDebug>
{
public:
	
	BEGIN_COM_MAP(CDebug)
		COM_INTERFACE_ENTRY(IObjectDebug)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP( )

	DECLARE_NO_REGISTRY()
	DECLARE_CLASSFACTORY()
	DECLARE_PROTECT_FINAL_CONSTRUCT()

	STDMETHODIMP alert(VARIANT vs )
	{
		CComVariant vd;
		if ( SUCCEEDED(VariantChangeType( &vd, &vs, VARIANT_NOVALUEPROP, VT_BSTR )))		
			::MessageBox( NULL, vd.bstrVal, L"msgbox", MB_OK );
		return S_OK;
	}
	STDMETHODIMP print(VARIANT vs )
	{
		CComVariant vd;
		if ( SUCCEEDED(VariantChangeType( &vd, &vs, VARIANT_NOVALUEPROP, VT_BSTR ))	)
			::_tprintf(vd.bstrVal);

		return S_OK;
	}

	static LPCWSTR name(){ return L"debug"; }
};


class __declspec(uuid("{256D0123-575B-4A66-BA6A-35D51A0D6A7E}")) BrokerServer
	: public ATL::CAtlDllModuleT<BrokerServer>
{
public:
	DECLARE_LIBID(__uuidof(BrokerServer))

};

static BrokerServer m;