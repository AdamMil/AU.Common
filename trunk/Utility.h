// Utility.h : Declaration of the CUtility

#pragma once
#include "resource.h"       // main symbols

#include "Common.h"


// CUtility

class ATL_NO_VTABLE CUtility : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CUtility, &CLSID_Utility>,
	public IDispatchImpl<IUtility, &IID_IUtility, &LIBID_AU_CommonLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CUtility();
 ~CUtility();

  DECLARE_REGISTRY_RESOURCEID(IDR_UTILITY)
  DECLARE_NOT_AGGREGATABLE(CUtility)
  BEGIN_COM_MAP(CUtility)
	  COM_INTERFACE_ENTRY(IUtility)
	  COM_INTERFACE_ENTRY(IDispatch)
  END_COM_MAP()
	DECLARE_PROTECT_FINAL_CONSTRUCT()

public:
  STDMETHODIMP StringToBin(/*[in]*/ BSTR sStr, /*[in,defaultvalue(-1)]*/ VARIANT_BOOL b8bit, /*[out,retval]*/ VARIANT *pvOut);
  STDMETHODIMP BinToString(/*[in]*/ VARIANT vBin, /*[in,defaultvalue(-1)]*/ VARIANT_BOOL b8bit, /*[out,retval]*/ BSTR *psRet);
  STDMETHODIMP HexToBin(/*[in]*/ BSTR sHex, /*[out,retval]*/ VARIANT *pvOut);
  STDMETHODIMP BinToHex(/*[in]*/ VARIANT vBin, /*[in,defaultvalue(-1)]*/ VARIANT_BOOL b8bit, /*[out,retval]*/ BSTR *psHex);
  STDMETHODIMP EncodeB64(/*[in]*/ VARIANT vIn, /*[in,defaultvalue(-1)]*/ VARIANT_BOOL b8bit, /*[out,retval]*/ VARIANT *pvOut);
  STDMETHODIMP DecodeB64(/*[in]*/ VARIANT vIn, /*[in,defaultvalue(-1)]*/ VARIANT_BOOL b8bit, /*[out,retval]*/ VARIANT *pvOut);
  STDMETHODIMP HashSHA1(/*[in]*/ VARIANT vIn, /*[out,retval]*/ VARIANT *pvOut);
  STDMETHODIMP CheckSHA1(/*[in]*/ VARIANT vIn, /*[in]*/ VARIANT vHash, /*[out,retval]*/ VARIANT_BOOL *pbOK);
  STDMETHODIMP VarType(/*[in]*/ VARIANT vIn, /*[out,retval]*/ short *pRet);
  STDMETHODIMP GetTickCount(/*[out,retval]*/ long *pnMS);
  STDMETHODIMP LogonUser(/*[in]*/ BSTR sUser, /*[in]*/ BSTR sPass, /*[in,defaultvalue("")]*/ BSTR sDomain);
  STDMETHODIMP RevertLogin();

protected:
  HANDLE m_hToken;
};

OBJECT_ENTRY_AUTO(__uuidof(Utility), CUtility)
