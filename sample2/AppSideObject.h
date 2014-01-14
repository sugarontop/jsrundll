#pragma once

#include "sample2_h.h"
#include "IDispatchImpl2.h"

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