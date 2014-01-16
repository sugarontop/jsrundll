#include "stdafx.h"
#include "jsrundll.h"

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
	
	
		JSRun2 run = (JSRun2)::GetProcAddress( h, "JSRun2" );

		run( cxt, L"function alert(s){ debug.alert(s + ' jsrundll'); }" ); // see object_debug.h
		run( cxt, L"alert( 'hello world' );" );		

		JSClose close = (JSClose)::GetProcAddress( h, "JSClose" );
		close( cxt );
	}
	
	FreeLibrary( h );

	return 0;
}

