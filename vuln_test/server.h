// server.h : Declaration of the Cserver

#ifndef __SERVER_H_
#define __SERVER_H_

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// Cserver
class ATL_NO_VTABLE Cserver : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<Cserver, &CLSID_server>,
	public IConnectionPointContainerImpl<Cserver>,
	public IDispatchImpl<Iserver, &IID_Iserver, &LIBID_VULNLib>
{
public:
	Cserver()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_SERVER)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(Cserver)
	COM_INTERFACE_ENTRY(Iserver)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IConnectionPointContainer)
END_COM_MAP()
BEGIN_CONNECTION_POINT_MAP(Cserver)
END_CONNECTION_POINT_MAP()


// Iserver
public:
	STDMETHOD(HeapCorruption)(/*[in]*/ BSTR strIn, /*[in]*/ long bufSize, /*[out,retval]*/ long *retVal);
	STDMETHOD(Method4)(/*[in]*/ BSTR sPath, /*[in]*/ BSTR msg , /*[out,retval]*/ long *retVal);
	STDMETHOD(Method3)(/*[in]*/ VARIANT, /*[out,retval]*/ long *retVal);
	STDMETHOD(Method2)(/*[in]*/ long, /*[out,retval]*/ long *retVal);
	STDMETHOD(Method1)(/*[in]*/ BSTR , /*[out,retval]*/ long *retVal);
};

#endif //__SERVER_H_
