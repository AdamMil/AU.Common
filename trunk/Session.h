// Session.h : Declaration of the CAUSession

#pragma once
#include "resource.h"       // main symbols

#include "Common.h"


// CAUSession

class ATL_NO_VTABLE CAUSession : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CAUSession, &CLSID_AUSession>,
	public IDispatchImpl<IAUSession, &IID_IAUSession, &LIBID_AU_CommonLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
  DECLARE_REGISTRY_RESOURCEID(IDR_SESSION)
  DECLARE_NOT_AGGREGATABLE(CAUSession)
  BEGIN_COM_MAP(CAUSession)
	  COM_INTERFACE_ENTRY(IAUSession)
	  COM_INTERFACE_ENTRY(IDispatch)
	  COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
  END_COM_MAP()
	DECLARE_PROTECT_FINAL_CONSTRUCT()
	DECLARE_GET_CONTROLLING_UNKNOWN()

	HRESULT FinalConstruct() { return CoCreateFreeThreadedMarshaler(GetControllingUnknown(), &m_pUnkMarshaler.p); }
	void FinalRelease() { m_pUnkMarshaler.Release(); }

	CComPtr<IUnknown> m_pUnkMarshaler;
public:
  STDMETHODIMP get_Item(/*[in]*/ BSTR sKey, /*[out,retval]*/ VARIANT *pvOut);
  STDMETHODIMP put_Item(/*[in]*/ BSTR sKey, /*[in]*/ VARIANT vData);
  STDMETHODIMP get_SessionID(/*[out,retval]*/ BSTR *psID);
  STDMETHODIMP get_Section(/*[out,retval]*/ BSTR *psSect);
  STDMETHODIMP put_Section(/*[in]*/ BSTR sSect);
  STDMETHODIMP get_Flags(/*[out,retval]*/ long *pnFlags);
  STDMETHODIMP put_Flags(/*[in]*/ long nFlags);
  STDMETHODIMP get_Timeout(/*[out,retval]*/ long *pnSecs);
  STDMETHODIMP put_Timeout(/*[in]*/ long nSecs);
  STDMETHODIMP get_IsManaged(/*[out,retval]*/ VARIANT_BOOL *pbManaged);
  STDMETHODIMP get_UsesDB(/*[out,retval]*/ VARIANT_BOOL *pbDB);
  STDMETHODIMP put_UsesDB(/*[in]*/ VARIANT_BOOL bDB);

  STDMETHODIMP Clear(/*[in]*/ BSTR sSect);
  STDMETHODIMP ClearAll();
  STDMETHODIMP Persist(/*[out,retval]*/ BSTR *psID);
  STDMETHODIMP Revert();
  STDMETHODIMP Save();
  STDMETHODIMP Delete();
};

OBJECT_ENTRY_AUTO(__uuidof(AUSession), CAUSession)
