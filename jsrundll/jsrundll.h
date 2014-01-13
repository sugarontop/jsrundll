#pragma once

extern "C" {

typedef int JSCONTEXT;

struct AppInfo
{
	DWORD size;
	WCHAR appname[32];
	WCHAR appguid[76];
	DWORD appversion;
	WCHAR authkey[32];
	WCHAR reg_current_user[128];
};

#ifdef JSRUN_CLIENT

	// Initialize
	typedef JSCONTEXT  (WINAPI *JSInitialiaze)( AppInfo* info, int version_number );

	// Set one dispatch static object 
	typedef JSCONTEXT (WINAPI *JSSetStaticObject)( JSCONTEXT cxt, LPCWSTR name, IDispatch* StaticObjects );

	// Get one dispatch static object 
	typedef bool (WINAPI *JSGetStatic)( JSCONTEXT cxt, LPCWSTR nm, IDispatch** pout );

	// Set some dispatch static objects
	typedef JSCONTEXT (WINAPI *JSSetStaticObjects)( JSCONTEXT cxt, LPOLESTR* StaticObjectName, IDispatch** StaticObjects, int cStatinObject );

	// Excecute program
	typedef bool (WINAPI *JSRun)( JSCONTEXT cxt, BSTR src );


	typedef bool (WINAPI *JSRun2)( JSCONTEXT cxt, LPCWSTR src );

	// GetErrorMessage
	typedef BSTR (WINAPI *JSErrorMsg)(JSCONTEXT cxt);

	// Close.
	typedef void (WINAPI *JSClose)(JSCONTEXT cxt);

#endif

};
