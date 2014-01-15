#include "stdafx.h"
#include "jsrundll.h"
#include "AppSideObject.h"
#include "InvokeHelper.h"

void debug_typecheck();

int _tmain(int argc, _TCHAR* argv[])
{
	CoInitialize(0);

	_tsetlocale ( 0, L"" );
	_CrtSetDbgFlag ( _CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF );

	debug_typecheck();

	HMODULE h = LoadLibrary( L"jsrun.dll" );
	{
		JSInitialiaze jsini = (JSInitialiaze)::GetProcAddress( h, "JSInitialiaze" );
		JSRun run = (JSRun)::GetProcAddress( h, "JSRun" );
		JSClose close = (JSClose)::GetProcAddress( h, "JSClose" );
		JSSetStaticObjects objadd = (JSSetStaticObjects)::GetProcAddress( h, "JSSetStaticObjects" );


		AppInfo info;
		info.size = sizeof(info);
		lstrcpy(info.appname, L"sample" );	
		lstrcpy(info.appguid, L"This is test." );
		info.appversion = 1;

		JSCONTEXT cxt = jsini( &info, JSCRIPT_ENGINE_VERSION );
				
		CComBSTR src;
		
		// event handler test
		{
			CComPtr<IDispatch> d;
		
			CHelper::CreateInstance( &d );
			LPOLESTR nm[]= { L"app" };
			IDispatch* pa[] = { d.p };

			objadd( cxt, nm, pa, 1 );
		
			src = L"function f1(a){ return a * 10;} ";
			src += L"function f2(a){ return a * 20;} ";
			src += L"app.EventHandler = f1; ";
			src += L"app.EventTest();";
				
			run( cxt, src ); // "result = 1000"

			src = L"app.EventHandler = f2; ";
			src += L"app.EventTest();";

			run( cxt, src );// "result = 2000"
		}
								
		close( cxt );
	}
	
	FreeLibrary( h );
	return 0;
} 

STDMETHODIMP CHelper::Hello()
{			
	_tprintf( L"Application side com." );
	return S_OK;
}
STDMETHODIMP CHelper::put_EventHandler( VARIANT js_function )
{
	eventhandler_ = js_function;
	return S_OK;
}
STDMETHODIMP CHelper::EventTest()
{			
	_ASSERT( eventhandler_.vt == VT_DISPATCH );

	CDispatchHelper c(eventhandler_);

	CComVariant rval;
	CComVariant vpm1 = 100;		// test value.

	c.Method1( (DISPID)0, rval, vpm1 ); // Unknown function is DISPID=0

	_tprintf( L"result = %d\n", rval.intVal );
			
	return S_OK;
}

////////////////////////////////////////////////////////////////
void debug_typecheck()
{
	TCHAR szFilePath[MAX_PATH];		
	DWORD dwFLen = ::GetModuleFileName(_AtlBaseModule.GetModuleInstance(), szFilePath, MAX_PATH);
	HRESULT hr = E_FAIL;
	
	if( dwFLen != 0 && dwFLen != MAX_PATH )
	{
		CComPtr<ITypeLib> pTypeLib;
		hr = LoadTypeLib(CT2W(szFilePath), &pTypeLib); // if fail, check idl,rc,tlb files.
	}

	ATLASSERT(SUCCEEDED(hr));
}