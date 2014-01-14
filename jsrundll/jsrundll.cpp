// jsrundll.cpp : DLL アプリケーション用にエクスポートされる関数を定義します。
//
#include "stdafx.h"
#include "jsrundll.h"
#include "cactivescriptsite.h"
#include "object_debug.h"

CComPtr<IActiveScript> pAS;
CComPtr<IActiveScriptParse> jsParse;

static CActiveScriptSite* gCAss;
static std::wstring gErrMsg;
static int version_number_g;
static JSCONTEXT gContextNumber = 0;
static std::shared_ptr<AppInfo> gAppInfo;

void Finalize();

template <class T> CComPtr<ITypeInfo> IDispatchImpl2<T>::pTypeInfo_;


bool Initial(int version_number)
{
	HRESULT hr = E_FAIL;
	version_number_g = -1;
	
	if ( version_number == 9 ) // JavaScript1.8.5
	{
		CLSID clsid;
		CComBSTR bs1 = L"{16D51579-A30B-4C8B-A276-0FF4DC41E755}"; // JScript9 chakra.
		hr = CLSIDFromString( bs1, &clsid );

		if ( hr == S_OK )
		{
			hr = pAS.CoCreateInstance( clsid, NULL );
			version_number_g = version_number;
		}
		else
			version_number = 8;
	}

	// try older version dll.
	if ( hr != S_OK || version_number < 9 )
	{
		CLSID clsid;
		hr = CLSIDFromProgID( L"JScript", &clsid ); // Default JScript5.7-5.8( 2012.10 ) 

		if ( hr == S_OK )
		{
			hr = pAS.CoCreateInstance( clsid, NULL );
			if ( hr == S_OK )
				version_number_g = 7;

			if ( hr == S_OK && version_number == 8 )
			{
				// In order to support new features after JScript 5.8(IE8,JavaScript1.5)
				CComPtr<IActiveScriptProperty> pActScriProp;
				if (SUCCEEDED(pAS->QueryInterface(&pActScriProp))) 
				{             
					CComVariant scriptLangVersion;             
					scriptLangVersion.vt = VT_I4;             
					scriptLangVersion.lVal = SCRIPTLANGUAGEVERSION_5_8;             
					pActScriProp->SetProperty(SCRIPTPROP_INVOKEVERSIONING, NULL, &scriptLangVersion);				

					version_number_g = 8;
				} 
			}
		}
	}	

	if ( hr == S_OK )
	{
		hr = pAS.QueryInterface( &jsParse );
		
		if ( hr != S_OK)
			version_number_g = -1;
	}

	return ( hr == S_OK );
}
JSCONTEXT gencontext( LPCWSTR key )
{
	gContextNumber = lstrlen( key );
	return gContextNumber;
}
int Initial2()
{
	gCAss = new CActiveScriptSite(pAS);

	
	if ( S_OK != pAS->SetScriptSite( gCAss) )
		return -1;
	if ( S_OK !=  jsParse->InitNew() )
		return -1;

	// Donot forget CActiveScriptSite::GetItemInfo.
	
	// static object in javascript.
	OLECHAR* ar[] = {	
		(OLECHAR*)CDebug::name(),
		//(OLECHAR*)CJSRunCom::name(),
		//(OLECHAR*)CMs::name(),
		//(OLECHAR*)CCryptoAES::name(), 
		//(OLECHAR*)CCryptoAES256::name(),
		//(OLECHAR*)CBinary::name(),		
		// (OLECHAR*)CMapCom::name(),
		//(OLECHAR*)CIoCom::name()
	};

	int cn = sizeof(ar)/sizeof(ar[0]);

	for( int i = 0; i < cn; i++ )
	{
		if ( S_OK !=  pAS->AddNamedItem( ar[i] , SCRIPTITEM_ISVISIBLE )) 
			return -1;
	}


	return 0;

}

DLLEXPORT JSCONTEXT  WINAPI JSInitialiaze(AppInfo* info, int version_number )
{
	gErrMsg.clear();

	_ASSERT( info->size == sizeof(AppInfo) );

	JSCONTEXT cxt = gencontext(info->authkey);
	//if ( !context_check(cxt)) return -1;


	if ( !Initial(version_number) )
	{
		gErrMsg = L"Initialize Error.";
		return -1;
	}
	if ( 0 == Initial2() )
	{
		#ifdef __Script_Debug__
		((CActiveScriptSite*)gCAss)->CreateScriptDebugger();
		#endif

		return cxt;
	}
	return -1; // error
		
}



DLLEXPORT bool WINAPI  JSRun( JSCONTEXT cxt, BSTR src1 )
{
	BSTR src = src1;
	if (  0 == ::SysStringLen( src ))
		return true;
	
	HRESULT hr;
	gErrMsg.clear();
	EXCEPINFO epi,epi2;
	ZeroMemory(&epi, sizeof(EXCEPINFO));
	ZeroMemory(&epi2, sizeof(EXCEPINFO));

	CComVariant ret;

	CActiveScriptSite* pASS = (CActiveScriptSite*)gCAss;
	pASS->ErrorReset();

	#ifdef __Script_Debug__
		((CActiveScriptSite*)gCAss)->CreateDocumentForDebugger(src);
	#endif

	DWORD cookie = pASS->pdwSourceContext;


	// ttps://technet.microsoft.com/ja-jp/subscriptions/securedownloads/tch4w30x(v=vs.94).aspx
	
	hr = jsParse->ParseScriptText( src, NULL,NULL,NULL,cookie,0,0 ,&ret, &epi2 );
	epi = pASS->Error();
		
	if ( epi.scode != 0 || hr != S_OK )
		goto err;

	hr = pAS->SetScriptState( SCRIPTSTATE_CONNECTED);


	epi = pASS->Error();

	HRESULT hr2 = pAS->SetScriptState( SCRIPTSTATE_DISCONNECTED);


	if ( epi.scode != 0 || hr != S_OK || hr2 != S_OK )
	{
		goto err;
	}
		
	return true;
err:
	if ( pASS->ErrLineNo() )
	{
		CString msg;
		msg.Format( L"ERROR row=%d col=%d, %s, %s \n", pASS->ErrLineNo()+1, pASS->ErrPos(), epi.bstrSource, epi.bstrDescription );
		gErrMsg = (LPCWSTR)msg;		

		
		::OutputDebugString( msg );
	}
	else
	{		
		CString msg;

		if ( hr == DISP_E_EXCEPTION )
		{
			msg = L"error throwing. DISP_E_EXCEPTION";
		}
		else
		{
			msg.Format( L"ERROR %s, %s, HR=%d\n",  epi.bstrSource, epi.bstrDescription, hr );
			
		}
		
		gErrMsg = (LPCWSTR)msg;		
		::OutputDebugString( msg );
	}


	return false;

}
DLLEXPORT bool WINAPI  JSRun2( JSCONTEXT cxt, LPCWSTR src )
{
	CComBSTR src1 = src;

	return JSRun(cxt, src1 );
}

void Finalize()
{
	if ( pAS.p == NULL ) 
		return;	

	#ifdef __Script_Debug__
	((CActiveScriptSite*)gCAss)->ReleaseScriptDebugger();
	#endif

	IActiveScriptGarbageCollector * gc = NULL;
	if (SUCCEEDED(pAS->QueryInterface(IID_IActiveScriptGarbageCollector, (void **)&gc)))
	{
		gc->CollectGarbage(SCRIPTGCTYPE_EXHAUSTIVE);
		gc->Release();
	}


	// do not forget FreeLibrary!
	pAS->SetScriptState( SCRIPTSTATE_DISCONNECTED);
	pAS->SetScriptState( SCRIPTSTATE_CLOSED);

	jsParse.Release();

	pAS->Close();
	pAS.Release();
}
DLLEXPORT BSTR WINAPI  JSErrorMsg(JSCONTEXT cxt)
{
	//if ( cxt > 0 )
	//	if ( !context_check(cxt)) return NULL;

	CComBSTR bs = gErrMsg.c_str();
	return bs;
}
DLLEXPORT void WINAPI  JSClose(JSCONTEXT cxt)
{
	//if ( !context_check(cxt)) return;
	Finalize();

	gAppInfo.reset();
}
DLLEXPORT JSCONTEXT  WINAPI JSSetStaticObjects( JSCONTEXT cxt, LPOLESTR* StaticObjectName, IDispatch** StaticObjects, int cStatinObject )
{
	//if ( !context_check(cxt)) return -1;

	gCAss->NewStaticObject( StaticObjectName, StaticObjects, cStatinObject );
	
	for( int i = 0; i < cStatinObject; i++ )
	{
		if ( S_OK !=  pAS->AddNamedItem( StaticObjectName[i] , SCRIPTITEM_GLOBALMEMBERS|SCRIPTITEM_ISVISIBLE )) // SCRIPTITEM_GLOBALMEMBERSでないとParseScriptTextで落ちる、GLOBALかどうかはサーバのことか？
		return -1;
	}
	return cxt;
}
DLLEXPORT JSCONTEXT  WINAPI JSSetStaticObject( JSCONTEXT cxt, LPCWSTR name, IDispatch* StaticObjects )
{
	IDispatch* disp[1];
	disp[0] = StaticObjects;

	LPOLESTR nms[1];
	nms[0] = (OLECHAR*)name;

	return JSSetStaticObjects( cxt, nms, disp, 1 );	
}

// Get AddNamedItem object.
// ex. LPOLESTR nm[] = { L"Debug" };
//  Array,Object,Function                
bool GetStaticObject( IActiveScript* pAS, LPOLESTR* nm, IDispatch** pout )
{
	HRESULT hr;
	CComPtr<IDispatch> pD;
	CComVariant vr;
	DISPID idMethod; 
	hr = pAS->GetScriptDispatch(NULL,&pD);
	_ASSERT( hr == S_OK );
	//LPOLESTR array_name[]={L"MyMessage"};
	DISPPARAMS no_params = {NULL,0,0,0};
	hr = pD->GetIDsOfNames(IID_NULL, nm, 1, LOCALE_USER_DEFAULT, &idMethod); 
	_ASSERT( hr == S_OK );
	hr = pD->Invoke(idMethod,IID_NULL,LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,&no_params,&vr,NULL,NULL);
	_ASSERT( hr == S_OK );

	*pout = vr.pdispVal;
	(*pout)->AddRef();
	return true;
}


DLLEXPORT bool WINAPI  JSGetStatic( JSCONTEXT cxt, LPCWSTR nm, IDispatch** pout )
{
	//if ( !context_check(cxt)) return false;

	LPOLESTR nms[1];
	nms[0] = (LPOLESTR)nm;
	return GetStaticObject( pAS, nms, pout );
}

	STDMETHODIMP CActiveScriptSite::GetItemInfo(LPCOLESTR pstrName,DWORD dwReturnMask,	IUnknown **ppiunkItem,ITypeInfo **ppti)
{
	if(dwReturnMask & SCRIPTINFO_IUNKNOWN)
	{
		if ( lstrcmpW(pstrName, CDebug::name() )==0)
		{
			CComPtr<IDispatch> disp;
			CDebug::CreateInstance( &disp ); // NOT REGISTORY CLASS

			*ppiunkItem=(IUnknown*)disp.p;
			disp.p->AddRef();

			{				
				CComPtr<IDispatch> arr;
				LPOLESTR nm[] = { L"Array" };
				bool bl =  GetStaticObject( pAS, nm, &arr );

				CComVariant var = arr;
				HRESULT hr;
				LPOLESTR nm2[] = { L"Value" };
				DISPID id; VARIANT v;EXCEPINFO ex; UINT er; DISPPARAMS pm = { NULL,0,0,0 };

				DISPID dispIDNamedArgs[1] = { DISPID_PROPERTYPUT };
				pm.cArgs = 1;
				pm.rgvarg = &var;
				pm.cNamedArgs = 1;
				pm.rgdispidNamedArgs = dispIDNamedArgs;

				hr = disp.p->GetIDsOfNames( IID_NULL, nm2, 1, 0, &id );
				hr = disp.p->Invoke( id, IID_NULL, 0, DISPATCH_PROPERTYPUT, &pm, &v, &ex, &er );
			}

			return S_OK;
		}
		else
		{
			std::map<std::wstring,IDispatch*>::iterator it = mapStaticObject_.find( pstrName );
			if ( it != mapStaticObject_.end())
			{
				IDispatch* p = mapStaticObject_[ pstrName ];
				*ppiunkItem=(IUnknown*)p;

				p->AddRef();
				return S_OK;
			}
		}
	}
	else if ( dwReturnMask & SCRIPTINFO_ITYPEINFO )
	{
		if ( lstrcmpW(pstrName, CDebug::name() )==0)
		{
			CComPtr<IDispatch> disp;
			CDebug::CreateInstance( &disp ); // NOT REGISTORY CLASS

			disp->GetTypeInfo( 0,LOCALE_USER_DEFAULT, ppti );
			return S_OK;
		}
	}
	return TYPE_E_ELEMENTNOTFOUND;
}