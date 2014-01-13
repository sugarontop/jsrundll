#pragma once

#include "jsrundll_h.h"
#include "IDispatchImpl2.h"
#include <DispEx.h>
using namespace ATL;

class __declspec(uuid("{15F47057-DE9B-4638-8B72-CABF72E93658}")) CDebug
	: public CComObjectRootEx<CComSingleThreadModel>
	, public CComCoClass<CDebug, &__uuidof(CDebug)>
	, public IDispatchImpl2<IObjectDebug>
{
public:
	
	CDebug()
	{
		version_ = L"0.1";
	}
	~CDebug()
	{
		int a = 0;
	}
	BEGIN_COM_MAP(CDebug)
		COM_INTERFACE_ENTRY(IObjectDebug)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP( )

	DECLARE_NO_REGISTRY()
	DECLARE_CLASSFACTORY()
	DECLARE_PROTECT_FINAL_CONSTRUCT()

	STDMETHODIMP alert(BSTR txt )
	{
		::MessageBox( NULL, txt, L"msgbox", MB_OK );
		return S_OK;
	}

/*
	STDMETHODIMP Hello(void) throw()
	{
		_tprintf(_TEXT("Hello, World!\n"));
		return S_OK;
	}

	STDMETHODIMP trace( VARIANT v )
	{
		CComVariant vs;
		if ( S_OK == vs.ChangeType( VT_BSTR, &v ) )
		{
			::OutputDebugString( vs.bstrVal );
			::OutputDebugString( L"\n" );
		}
		else
		{
			CString s;
			s.Format( L"variant bstr change err: vt=%d\n", v.vt );
			::OutputDebugString( s );
		}		
		return S_OK;
	}
	STDMETHODIMP vt( VARIANT v, UINT* ret )
	{
		*ret = v.vt;
		return S_OK;
	}
	STDMETHODIMP print( VARIANT v )
	{
		CComVariant vs;
		if ( S_OK == vs.ChangeType( VT_BSTR, &v ) )
		{
			_tprintf( vs.bstrVal );
			_tprintf( L"\n" );
		}
		else
		{
			CString s;
			s.Format( L"variant bstr change err: vt=%d\n", v.vt );
			_tprintf( s );
		}		
		return S_OK;
	}

	STDMETHODIMP put_Version( VARIANT v )
	{
		return S_OK;
	}

	STDMETHODIMP get_Version( VARIANT* v )
	{		
		VariantCopy( v, &version_ );
		return S_OK;
	}

	STDMETHODIMP SetValue( BSTR key, VARIANT val )
	{
		map_[ key ] = val;
		return S_OK;
	}
	STDMETHODIMP GetValue( BSTR key, VARIANT* val )
	{
		if ( map_.find( key ) != map_.end())
		{
			// VariantCopyは内部でp->AddRef()あり
			// VariantClearは内部でp->Release()あり

			VariantCopy( val, &(map_[ key ]));
		}
		return S_OK;

	}

	STDMETHODIMP put_Value( VARIANT val )
	{
		// jsrun.cpp内でArrayを割りあてている。
		value_ = val;
		return S_OK;
	}
	STDMETHODIMP get_Value( VARIANT* val )
	{
		VariantCopy( val, &value_ );
		return S_OK;

	}
	 
	STDMETHODIMP traceEnumMethod( VARIANT v )
	{
		CComPtr<IDispatchEx> dx;
		if ( v.vt == VT_DISPATCH && S_OK == v.pdispVal->QueryInterface( &dx ))
		{
			DISPID dispid;
			HRESULT hr = dx->GetNextDispID(fdexEnumAll, DISPID_STARTENUM, &dispid);
			while( hr != S_FALSE)
			{
				CComBSTR bs;
				if ( S_OK != dx->GetMemberName( dispid, &bs ) )
					break;

				::OutputDebugString( bs );
				::OutputDebugString( L"\n" );
				hr = dx->GetNextDispID(fdexEnumAll, dispid, &dispid);
			}
		}		
		return S_OK;
	}
	STDMETHODIMP LoadLib( BSTR jsscript, VARIANT* ret );
	STDMETHODIMP Split(BSTR txt, BSTR chr, VARIANT* ret );
	


	typedef void (*EventHandler1)(BSTR txt );

	STDMETHODIMP EventFire( BSTR txt )
	{
		if ( eventhandler_.vt == VT_INT )
		{
			// 型チエックできない
			EventHandler1 ev = (EventHandler1)eventhandler_.intVal;
			ev( txt );
		}

		return S_OK;
	}
	STDMETHODIMP put_EventFunc(  VARIANT v  )
	{
		eventhandler_ = v;
		return S_OK;
	}
*/
	static LPCWSTR name(){ return L"debug"; }

	CComVariant eventhandler_;
	CComVariant version_;
	CComVariant value_;

	std::map<std::wstring, CComVariant> map_;

};
class __declspec(uuid("{256D0123-575B-4A66-BA6A-35D51A0D6A7E}")) BrokerServer
	: public ATL::CAtlDllModuleT<BrokerServer>
{
public:
	DECLARE_LIBID(__uuidof(BrokerServer))

};

static BrokerServer m;