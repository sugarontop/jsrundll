#pragma once


#include "sample2_h.h"

using namespace ATL;
template<class T> class IDispatchImpl2 : public T
{
public:

	virtual HRESULT __stdcall GetTypeInfoCount(unsigned int FAR* pctinfo) throw()
	{
		if ( !pctinfo) {
			return E_POINTER;
		}

		HRESULT hr = EnsureTI();

		*pctinfo = SUCCEEDED(hr) ? 1 : 0;

		return S_OK;
	}

	virtual HRESULT __stdcall GetTypeInfo(unsigned int iTInfo, LCID lcid,
		ITypeInfo FAR* FAR* ppTInfo) throw()
	{
		HRESULT hr = EnsureTI();
		if (SUCCEEDED(hr)) {
			return pTypeInfo_.CopyTo(ppTInfo);
		}
		return TYPE_E_ELEMENTNOTFOUND;
	}

	virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, OLECHAR FAR* FAR* rgszNames,
		unsigned int cNames, LCID lcid, DISPID FAR* rgDispId) throw()
	{
		HRESULT hr = EnsureTI();
		if (SUCCEEDED(hr)) {
			hr = pTypeInfo_->GetIDsOfNames(rgszNames, cNames, rgDispId);
		}
		return hr;
	}

	virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
		WORD wFlags, DISPPARAMS FAR* pDispParams, VARIANT FAR* pVarResult,
		EXCEPINFO FAR* pExcepInfo, unsigned int FAR* puArgErr) throw()
	{
		HRESULT hr = EnsureTI();
		if (SUCCEEDED(hr)) {
			hr = pTypeInfo_->Invoke(this, dispIdMember, wFlags,
				pDispParams, pVarResult, pExcepInfo, puArgErr);
		}
		return hr;
	}

	static HRESULT EnsureTI() throw()
	{
		if (pTypeInfo_) 
		{
			return S_OK;
		}
		HRESULT hr = E_FAIL;
		TCHAR szFilePath[MAX_PATH];
		DWORD dwFLen = ::GetModuleFileName(	_AtlBaseModule.GetModuleInstance(), szFilePath, MAX_PATH);
		if( dwFLen != 0 && dwFLen != MAX_PATH ) 
		{
			CComPtr<ITypeLib> pTypeLib;
			hr = LoadTypeLib(CT2W(szFilePath), &pTypeLib);
			ATLASSERT(SUCCEEDED(hr));
			if (SUCCEEDED(hr)) 
			{
				const GUID guid = __uuidof(T);
				hr = pTypeLib->GetTypeInfoOfGuid(guid, &pTypeInfo_);
				ATLASSERT(SUCCEEDED(hr));
			}
		}
		return hr;
	}

	static CComPtr<ITypeInfo> pTypeInfo_;
};


class __declspec(uuid("{9A4E4CD1-C9F6-490D-9A90-3F783E20DCA6}")) CHelper
	: public CComObjectRootEx<CComSingleThreadModel>
	, public CComCoClass<CHelper, &__uuidof(CHelper)>
	, public IDispatchImpl2<IObjectAppHelper>
{
public:

	BEGIN_COM_MAP(CHelper)
		COM_INTERFACE_ENTRY(IObjectAppHelper)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP( )

	DECLARE_NO_REGISTRY()
	DECLARE_CLASSFACTORY()
	DECLARE_PROTECT_FINAL_CONSTRUCT()

	STDMETHODIMP Hello();
};



class __declspec(uuid("{3E46DC0A-DC9E-433c-89AB-9E271E1DBA0A}")) BrokerAppServer
	: public ATL::CAtlExeModuleT<BrokerAppServer>
{
public:
	DECLARE_LIBID(__uuidof(BrokerAppServer))

};


static BrokerAppServer m;
template <class T> CComPtr<ITypeInfo> IDispatchImpl2<T>::pTypeInfo_;