/* 
   This file is part of the AU.Common library, a set of ActiveX
   controls and C++ classes used to aid in COM and Web development.
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

/* ~(MODULES::DB, c'DB
  The DB object (ProgID AU.Common.DB) provides a nice wrapper around the MS ADO interface.
  The object uses the `Config' class to retrieve its connection information from the
  default configuration file using the Section/Key paradigm. However, it also supports a
  using a direct connection string. The DB object is thread safe and ASP safe, and can
  be used by multiple threads. However, if a thread is going to perform a batch of
  operations (such as setting the Lock/Cursor types and then doing an Execute), it should
  use `LockDB' and `UnlockDB' to lock the DB object for more than a single call. The DB
  object will attempt to keep itself closed for as long as possible. The DB object uses
  the "both" threading model. However, you cannot store ADO objects in session variables
  unless you have switched ADO into free threaded mode.
)~ */

class ATL_NO_VTABLE CDB : 
  public CComObjectRootEx<CComMultiThreadModel>,
  public CComCoClass<CDB, &CLSID_DB>,
  public IDispatchImpl<IDB, &IID_IDB, &LIBID_AU_CommonLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
  CDB();
 ~CDB();
   
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
  STDMETHODIMP get_ConnSection(/*[out, retval]*/ BSTR *psSect);
  STDMETHODIMP put_ConnSection(/*[in]*/ BSTR sSect);
  STDMETHODIMP get_ConnKey(/*[out, retval]*/ BSTR *psKey);
  STDMETHODIMP put_ConnKey(/*[in]*/ BSTR sKey);
  STDMETHODIMP get_ConnString(/*[out, retval]*/ BSTR *psStr);
  STDMETHODIMP put_ConnString(/*[in]*/ BSTR sStr);
  STDMETHODIMP get_CursorType(/*[out, retval]*/ int *pnType);
  STDMETHODIMP put_CursorType(/*[in]*/ int nType);
  STDMETHODIMP get_LockType(/*[out, retval]*/ int *pnType);
  STDMETHODIMP put_LockType(/*[in]*/ int nType);
  STDMETHODIMP get_Timeout(/*[out, retval]*/ long *pnTimeout);
  STDMETHODIMP put_Timeout(/*[in]*/ long nTimeout);
  STDMETHODIMP get_Connection(/*[out, retval]*/ ADOConnection **ppConn);
  STDMETHODIMP get_Command(/*[out, retval]*/ ADOCommand **ppCmd);
  STDMETHODIMP get_IsOpen(/*[out, retval]*/ VARIANT_BOOL *pbOpen);

  STDMETHODIMP Open();
  STDMETHODIMP Close();
  STDMETHODIMP NewCommand(/*[out, retval]*/ ADOCommand **ppCmd);
  STDMETHODIMP LockDB();
  STDMETHODIMP UnlockDB();
  STDMETHODIMP Execute(/*[in]*/ BSTR sSQL, /*[in]*/ SAFEARRAY(VARIANT) *aVals, /*[out,retval]*/ ADORecordset **ppRS);
  STDMETHODIMP ExecuteO(/*[in]*/ BSTR sSQL, /*[in]*/ BSTR sParms, /*[in]*/ SAFEARRAY(VARIANT) *aVals, /*[out,retval]*/ ADORecordset **ppRS);
  STDMETHODIMP ExecuteVal(/*[in]*/ BSTR sSQL, /*[in]*/ SAFEARRAY(VARIANT) *aVals, /*[out,retval]*/ VARIANT *pRet);
  STDMETHODIMP ExecuteValO(/*[in]*/ BSTR sSQL, /*[in]*/ BSTR sParms, /*[in]*/ SAFEARRAY(VARIANT) *aVals, /*[out,retval]*/ VARIANT *pvOut);
  STDMETHODIMP ExecuteNR(/*[in]*/ BSTR sSQL, /*[in]*/ SAFEARRAY(VARIANT) *aVals);
  STDMETHODIMP ExecuteONR(/*[in]*/ BSTR sSQL, /*[in]*/ BSTR sParms, /*[in]*/ SAFEARRAY(VARIANT) *aVals);
  STDMETHODIMP Output(/*[in]*/ BSTR sParam, /*[out,retval]*/ VARIANT *pvOut);
  
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
  BSTR m_ConnStr;
  UA4  m_nTimeout;
  int  m_CursorType, m_LockType;
  bool m_bInit;
  
  static const DataTypeEnum CDB::m_dbTypes[VT_UINT];

  friend class CDBMan;
};

OBJECT_ENTRY_AUTO(__uuidof(DB), CDB)
