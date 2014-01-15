#include "stdafx.h"
#include "jsrundll.h"
#include "AppSideObject.h"

int _tmain(int argc, _TCHAR* argv[])
{
	CoInitialize(0);

	_tsetlocale ( 0, L"" );
	_CrtSetDbgFlag ( _CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF );

	HMODULE h = LoadLibrary( L"jsrun.dll" );
	{
		JSInitialiaze jsini = (JSInitialiaze)::GetProcAddress( h, "JSInitialiaze" );

		AppInfo info;
		info.size = sizeof(info);
		lstrcpy(info.appname, L"sample" );	
		lstrcpy(info.appguid, L"This is test." );
		info.appversion = 1;

		JSCONTEXT cxt = jsini( &info, JSCRIPT_ENGINE_VERSION );

		JSSetStaticObjects objadd = (JSSetStaticObjects)::GetProcAddress( h, "JSSetStaticObjects" );

		CComPtr<IDispatch> d;
		HRESULT hr = CHelper::CreateInstance( &d );
		LPOLESTR nm[]= { L"app" };
		IDispatch* pa[] = { d.p };

		objadd( cxt, nm, pa, 1 );

		JSRun2 run = (JSRun2)::GetProcAddress( h, "JSRun2" );

		run( cxt, L"app.Hello();" );	// "Application side com."


		JSClose close = (JSClose)::GetProcAddress( h, "JSClose" );
		close( cxt );
	}
	
	FreeLibrary( h );
}

STDMETHODIMP CHelper::Hello(void)
{			
	_tprintf( L"Application side com." );
	return S_OK;
}