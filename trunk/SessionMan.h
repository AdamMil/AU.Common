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

// SessionMan.h : Declaration of the CSessionMan

#pragma once
#include "resource.h"       // main symbols

#include "Common.h"


// CSessionMan

class ATL_NO_VTABLE CSessionMan : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CSessionMan, &CLSID_SessionMan>,
	public IDispatchImpl<ISessionMan, &IID_ISessionMan, &LIBID_AU_CommonLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
  DECLARE_REGISTRY_RESOURCEID(IDR_SESSIONMAN)
  DECLARE_NOT_AGGREGATABLE(CSessionMan)
  BEGIN_COM_MAP(CSessionMan)
	  COM_INTERFACE_ENTRY(ISessionMan)
	  COM_INTERFACE_ENTRY(IDispatch)
	  COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
  END_COM_MAP()
	DECLARE_PROTECT_FINAL_CONSTRUCT()
	DECLARE_GET_CONTROLLING_UNKNOWN()

	HRESULT FinalConstruct() { return CoCreateFreeThreadedMarshaler(GetControllingUnknown(), &m_pUnkMarshaler.p); }
	void FinalRelease() { m_pUnkMarshaler.Release(); }
	CComPtr<IUnknown> m_pUnkMarshaler;

public:
  STDMETHODIMP get_NumSessions(/*[out,retval]*/ long *pnSess);
  STDMETHODIMP get_Session(/*[in]*/ long nIndex, /*[out,retval]*/ IAUSession **ppSess);
  STDMETHODIMP get_Current(/*[out,retval]*/ IAUSession **ppSess);
  STDMETHODIMP Init(/*[in,defaultvalue("")]*/ BSTR sConfigKey, /*[in,defaultvalue(0)]*/ long nFlags);
  STDMETHODIMP CreateSession(/*[in,defaultvalue(0)]*/ VARIANT_BOOL bPersist, /*[out,retval]*/ IAUSession **ppSess);
  STDMETHODIMP LoadSession(/*[in]*/ BSTR sID, /*[out,retval]*/ IAUSession **ppSess);
  STDMETHODIMP DropSession(/*[in]*/ BSTR sID);
  STDMETHODIMP LockMan();
  STDMETHODIMP UnlockMan();
};

OBJECT_ENTRY_AUTO(__uuidof(SessionMan), CSessionMan)
