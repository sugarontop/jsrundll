#include "stdafx.h"
#include "invokehelper.h"

// DISPID id; VARIANT v;EXCEPINFO ex; UINT er; DISPPARAMS pm = { NULL,0,0,0 };

HRESULT InvokeSetProperty( IDispatch* disp, DISPID id, VARIANT& val )
{
	VARIANT v;EXCEPINFO ex; UINT er; DISPPARAMS pm = { NULL,0,0,0 };
	DISPID dispIDNamedArgs[1] = { DISPID_PROPERTYPUT };

	pm.rgvarg = &val;
	pm.cArgs =1;
	pm.cNamedArgs = 1;
	pm.rgdispidNamedArgs = dispIDNamedArgs;

	return disp->Invoke( id, IID_NULL,LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &pm, &v, &ex, &er );
}
HRESULT InvokeGetProperty( IDispatch* disp, DISPID id, VARIANT& val )
{
	EXCEPINFO ex; UINT er; DISPPARAMS pm = { NULL,0,0,0 };
	return disp->Invoke( id, IID_NULL,LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &pm, &val, &ex, &er );
}
HRESULT InvokeMethod0( IDispatch* disp, DISPID id, VARIANT& ret )
{
	EXCEPINFO ex; UINT er; DISPPARAMS pm = { NULL,0,0,0 };
	return disp->Invoke( id, IID_NULL,LOCALE_USER_DEFAULT, DISPATCH_METHOD, &pm, &ret, &ex, &er );
}
HRESULT InvokeMethod1( IDispatch* disp, DISPID id, VARIANT& ret, VARIANT& pm1 )
{
	EXCEPINFO ex; UINT er; DISPPARAMS pm = { NULL,0,0,0 };
	pm.rgvarg = &pm1;
	pm.cArgs =1;
	return disp->Invoke( id, IID_NULL,LOCALE_USER_DEFAULT, DISPATCH_METHOD, &pm, &ret, &ex, &er );
}
HRESULT InvokeMethod2( IDispatch* disp, DISPID id, VARIANT& ret, VARIANT& pm1,VARIANT& pm2 )
{
	EXCEPINFO ex; UINT er; DISPPARAMS pm = { NULL,0,0,0 };
	VARIANTARG vm[2];

	for( int i = 0; i < 2; i++ ) VariantInit( &vm[i] );

	vm[0] = pm2;
	vm[1] = pm1;
	
	pm.rgvarg = vm;
	pm.cArgs =2;
	return disp->Invoke( id, IID_NULL,LOCALE_USER_DEFAULT, DISPATCH_METHOD, &pm, &ret, &ex, &er );
}
HRESULT InvokeMethod3( IDispatch* disp, DISPID id, VARIANT& ret, VARIANT& pm1,VARIANT& pm2,VARIANT& pm3 )
{
	EXCEPINFO ex; UINT er; DISPPARAMS pm = { NULL,0,0,0 };
	VARIANTARG vm[3];
	for( int i = 0; i < 3; i++ ) VariantInit( &vm[i] );
	vm[0] = pm3;
	vm[1] = pm2;
	vm[2] = pm1;
	
	pm.rgvarg = vm;
	pm.cArgs =3;
	return disp->Invoke( id, IID_NULL,LOCALE_USER_DEFAULT, DISPATCH_METHOD, &pm, &ret, &ex, &er );
}
HRESULT InvokeMethod4( IDispatch* disp, DISPID id, VARIANT& ret, VARIANT& pm1,VARIANT& pm2,VARIANT& pm3, VARIANT& pm4 )
{
	EXCEPINFO ex; UINT er; DISPPARAMS pm = { NULL,0,0,0 };
	VARIANTARG vm[4];
	for( int i = 0; i < 4; i++ ) VariantInit( &vm[i] );
	vm[0] = pm4;
	vm[1] = pm3;
	vm[2] = pm2;
	vm[3] = pm1;
	
	pm.rgvarg = vm;
	pm.cArgs =4;
	return disp->Invoke( id, IID_NULL,LOCALE_USER_DEFAULT, DISPATCH_METHOD, &pm, &ret, &ex, &er );
}
HRESULT InvokeGetDISPID( IDispatch* disp, LPCWSTR nm, DISPID& id )
{
	LPOLESTR nma[1]; nma[0] = (LPOLESTR)nm;

	return disp->GetIDsOfNames( IID_NULL, nma, 1,LOCALE_USER_DEFAULT, &id );
}

HRESULT InvokeSetProperty( IDispatch* disp, LPCWSTR nm, VARIANT& val )
{
	DISPID id;
	HRESULT hr = InvokeGetDISPID( disp, nm, id );
	if ( SUCCEEDED(hr))
		hr = InvokeSetProperty( disp, id, val );
	return hr;
}
HRESULT InvokeGetProperty( IDispatch* disp, LPCWSTR nm, VARIANT& val )
{
	DISPID id;
	HRESULT hr = InvokeGetDISPID( disp, nm, id );
	if ( SUCCEEDED(hr))
		hr = InvokeGetProperty( disp, id, val );
	return hr;
}

HRESULT InvokeMethod0( IDispatch* disp, LPCWSTR nm, VARIANT& ret )
{
	DISPID id;
	HRESULT hr = InvokeGetDISPID( disp, nm, id );
	if ( SUCCEEDED(hr))
		hr = InvokeMethod0( disp, id, ret );
	return hr;
}
HRESULT InvokeMethod1( IDispatch* disp, LPCWSTR nm, VARIANT& ret, VARIANT& pm1 )
{
	DISPID id;
	HRESULT hr = InvokeGetDISPID( disp, nm, id );
	if ( SUCCEEDED(hr))
		hr = hr = InvokeMethod1( disp, id, ret,pm1 );
	return hr;
}
HRESULT InvokeMethod2( IDispatch* disp, LPCWSTR nm, VARIANT& ret, VARIANT& pm1,VARIANT& pm2 )
{
	DISPID id;
	HRESULT hr = InvokeGetDISPID( disp, nm, id );
	if ( SUCCEEDED(hr))
		hr = hr = InvokeMethod2( disp, id, ret,pm1,pm2 );
	return hr;
}
HRESULT InvokeMethod3( IDispatch* disp, LPCWSTR nm, VARIANT& ret, VARIANT& pm1,VARIANT& pm2,VARIANT& pm3 )
{
	DISPID id;
	HRESULT hr = InvokeGetDISPID( disp, nm, id );
	if ( SUCCEEDED(hr))
		hr = InvokeMethod3( disp, id, ret,pm1,pm2,pm3 );
	return hr;
}
HRESULT InvokeMethod4( IDispatch* disp, LPCWSTR nm, VARIANT& ret, VARIANT& pm1,VARIANT& pm2,VARIANT& pm3,VARIANT& pm4 )
{
	DISPID id;
	HRESULT hr = InvokeGetDISPID( disp, nm, id );
	if ( SUCCEEDED(hr))
		hr = InvokeMethod4( disp, id, ret,pm1,pm2,pm3,pm4 );
	return hr;
}
