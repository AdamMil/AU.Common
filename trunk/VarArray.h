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

// VarArray.h : Declaration of the CVarArray

#pragma once
#include "resource.h"       // main symbols

#include "Common.h"


// CVarArray

class ATL_NO_VTABLE CVarArray : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVarArray, &CLSID_VarArray>,
	public IDispatchImpl<IVarArray, &IID_IVarArray, &LIBID_AU_CommonLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
  CVarArray() { m_Var.v.vt=VT_NULL, m_Max=0; }
  DECLARE_REGISTRY_RESOURCEID(IDR_VARARRAY)
  DECLARE_NOT_AGGREGATABLE(CVarArray)
  BEGIN_COM_MAP(CVarArray)
	  COM_INTERFACE_ENTRY(IVarArray)
	  COM_INTERFACE_ENTRY(IDispatch)
  END_COM_MAP()
	DECLARE_PROTECT_FINAL_CONSTRUCT()

// IVarArray
public:
  STDMETHODIMP get_Item(/*[in]*/ long nIndex, /*[out,retval]*/ VARIANT *pvOut);
  STDMETHODIMP get_Size(/*[out,retval]*/ long *pnEls);
  STDMETHODIMP get_Capacity(/*[out,retval]*/ long *pnEls);
  STDMETHODIMP put_Capacity(/*[in]*/ long nEls);
  STDMETHODIMP get_Array(/*[out,retval]*/ VARIANT *pvOut);
  STDMETHODIMP Add(/*[in]*/ VARIANT vEl);
  
protected:
  AVAR m_Var;
  UA4  m_Max;
};

OBJECT_ENTRY_AUTO(__uuidof(VarArray), CVarArray)
