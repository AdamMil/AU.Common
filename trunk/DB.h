/* 
   This file is part of the AU.Common library, a set of ActiveX
   controls to aid in Web development.
   Copyright (C) 2002 Adam Milazzo

   This library is free software; you can redistribute it and/or
   modify it under the terms of the GNU Lesser General Public
   License as published by the Free Software Foundation; either
   version 2.1 of the License, or (at your option) any later version.

   This library is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
   Lesser General Public License for more details.

   You should have received a copy of the GNU Lesser General Public
   License along with this library; if not, write to the Free Software
   Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*/

// DB.h : Declaration of the CDB

#pragma once
#include "resource.h"       // main symbols

#include "Common.h"


// CDB

class ATL_NO_VTABLE CDB : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CDB, &CLSID_DB>,
	public IDispatchImpl<IDB, &IID_IDB, &LIBID_AU_CommonLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
  CDB();
  
  DECLARE_REGISTRY_RESOURCEID(IDR_DB)
  DECLARE_NOT_AGGREGATABLE(CDB)

  BEGIN_COM_MAP(CDB)
	  COM_INTERFACE_ENTRY(IDB)
	  COM_INTERFACE_ENTRY(IDispatch)
	  COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
  END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()
	DECLARE_GET_CONTROLLING_UNKNOWN()

	HRESULT FinalConstruct()
	{ return CoCreateFreeThreadedMarshaler(GetControllingUnknown(), &m_pUnkMarshaler.p);
	}

	void FinalRelease()
	{ m_pUnkMarshaler.Release();
	}

public:
  // IDB
  STDMETHODIMP get_ConnSection(/*[out, retval]*/ BSTR *pVal);
	STDMETHODIMP put_ConnSection(/*[in]*/ BSTR newVal);
	STDMETHODIMP get_ConnKey(/*[out, retval]*/ BSTR *pVal);
	STDMETHODIMP put_ConnKey(/*[in]*/ BSTR newVal);
	STDMETHODIMP get_CursorType(/*[out, retval]*/ int *pVal);
	STDMETHODIMP put_CursorType(/*[in]*/ int newVal);
	STDMETHODIMP get_LockType(/*[out, retval]*/ int *pVal);
	STDMETHODIMP put_LockType(/*[in]*/ int newVal);
	STDMETHODIMP get_Timeout(/*[out, retval]*/ long *pVal);
	STDMETHODIMP put_Timeout(/*[in]*/ long newVal);
	STDMETHODIMP get_Connection(/*[out, retval]*/ ADOConnection **pVal);
	STDMETHODIMP get_Command(/*[out, retval]*/ ADOCommand **pVal);
	STDMETHODIMP get_IsOpen(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHODIMP get_Output(/*[in]*/ BSTR sParam, /*[out,retval]*/ VARIANT *pvOut);

  STDMETHODIMP Open();
	STDMETHODIMP Close();
  STDMETHODIMP NewCommand(/*[out, retval]*/ ADOCommand **pVal);
  STDMETHODIMP LockDB();
  STDMETHODIMP UnlockDB();
  STDMETHODIMP Execute(/*[in]*/ BSTR sSQL, /*[in]*/ SAFEARRAY(VARIANT) *aVals, /*[out,retval]*/ ADORecordset **pRet);
  STDMETHODIMP ExecuteO(/*[in]*/ BSTR sSQL, /*[in]*/ BSTR sParms, /*[in]*/ SAFEARRAY(VARIANT) *aVals, /*[out,retval]*/ ADORecordset **pRet);
  STDMETHODIMP ExecuteVal(/*[in]*/ BSTR sSQL, /*[in]*/ SAFEARRAY(VARIANT) *aVals, /*[out,retval]*/ VARIANT *pRet);
	STDMETHODIMP ExecuteNR(/*[in]*/ BSTR sSQL, /*[in]*/ SAFEARRAY(VARIANT) *aVals);
	STDMETHODIMP ExecuteONR(/*[in]*/ BSTR sSQL, /*[in]*/ BSTR sParms, /*[in]*/ SAFEARRAY(VARIANT) *aVals);
	
protected:
  HRESULT Init();
  HRESULT StartExecute(BSTR sSQL);
  HRESULT DoExecute(ADORecordset **);
  void    ResetDefaults();
  HRESULT FillParams(SAFEARRAY(VARIANT) *aVals);
  HRESULT FillParamsO(BSTR sParms, SAFEARRAY(VARIANT) *aVals);

  AComPtr<ADOConnection> m_Conn;
  AComPtr<ADOCommand>    m_Cmd;
	CComPtr<IUnknown>      m_pUnkMarshaler;
  ASTR m_ConnSect, m_ConnKey;
  UA4  m_nTimeout;
  int  m_CursorType, m_LockType;
  bool m_bInit;
};

OBJECT_ENTRY_AUTO(__uuidof(DB), CDB)
