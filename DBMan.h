/* 
   This file is part of the AU.Common library, a set of ActiveX
   controls and C++ classes used to aid in COM and Web development.
   Copyright (C) 2002 Adam Milazzo - http://www.adammil.net

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

// DBMan.h : Declaration of the CDBMan

#pragma once
#include "resource.h"       // main symbols

#include "Common.h"

// CDBMan

/* ~(MODULES::DBMAN, c'DBMan
  The DBMan object (ProgID AU.Common.DBMan) implements database connection
  sharing. This is useful to reduce the number of connection open at a time,
  especially when those connections are known to not be used concurrently.
  An example is a function or ASP page that uses several COM objects that
  each need database connectivity. If each created their own connection,
  but the objects aren't accessed concurrently (which is generally the case),
  that page/function might use 5 database connections when it really could be
  using one. By using the DBMan object, those 5 connections would be pooled into
  a single one, increasing efficiency. Even if the DB objects are being used
  concurrently, it may not be a problem because DB is thread-safe (However, the
  locking on the DB object may cause a performance problem if too many threads are
  using a DB object concurrently).
  Additionally, since the DBMan object holds a reference to the database
  objects  and is Session-safe, it can be stored in the ASP session to prevent
  even the overhead of opening and closing the connection on each page hit.
  There are some additional things to be careful of when using DB objects that
  were created by DBMan, however. Since the objects are shared, it is not a
  good idea to alter their connection or to otherwise leave them in a state that
  could confuse other processes (for instance, setting the LockType and not
  resetting it [with Execute* or explicitly setting it back]). If the DB objects
  are going to be shared among multiple threads, you should follow the rules
  outlined in the `DB' object documentation. Additionally, you can't trust the
  DB object's `DB::ConnSection' and `DB::ConnKey' properties to be accurate. The
  `DB::ConnString' property should be accurate, though. The DBMan object uses
  the "both" threading model.
)~ */

class CDB;

class ATL_NO_VTABLE CDBMan : 
  public CComObjectRootEx<CComMultiThreadModel>,
  public CComCoClass<CDBMan, &CLSID_DBMan>,
  public IDispatchImpl<IDBMan, &IID_IDBMan, &LIBID_AU_CommonLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
  CDBMan();
 ~CDBMan();

  DECLARE_REGISTRY_RESOURCEID(IDR_DBMAN)
  DECLARE_NOT_AGGREGATABLE(CDBMan)
  BEGIN_COM_MAP(CDBMan)
    COM_INTERFACE_ENTRY(IDBMan)
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
  CComPtr<IUnknown> m_pUnkMarshaler;

public:
  // IDBMan
  STDMETHODIMP get_MaxSharing(/*[out,retval]*/ long *pnShares);
  STDMETHODIMP put_MaxSharing(/*[in]*/ long nShares);
  STDMETHODIMP CreateDB(/*[in,defaultvalue("")]*/ BSTR sKey, /*[in,defaultvalue("")]*/ BSTR sSect, /*[out,retval]*/ IDB **ppDB);
  STDMETHODIMP CreateDBRaw(/*[in]*/ BSTR sConnStr, /*[out,retval]*/ IDB **ppDB);

protected:
  class BSLess
  { public:
    bool operator()(const BSTR a, const BSTR b) const { return wcscmp(a, b)<0; }
  };
  typedef std::vector<CDB *> DBList;
  typedef std::map<BSTR, DBList *, BSLess> DBMap;
  DBMap m_Map;
  UA4   m_MaxShares;
};

OBJECT_ENTRY_AUTO(__uuidof(DBMan), CDBMan)
