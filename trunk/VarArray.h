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
/* ~(MODULES::VARARR, c'VarArray
  The VarArray object (ProgID AU.Common.VarArray) provides a simple method to create
  an array of VARIANTs. It is useful when building SQL queries from many optional
  parts from languages where you cannot pass an array directly to the Execute*() methods
  of the `DB' object. You can build an array incrementally with the VarArray object and
  assign it to the `DB::ParmArray' property. Then, the next Execute*() call will use that
  array instead of the one passed to Execute*(), allowing you to pass a variable number
  of parameters to the Execute*() calls from languages that otherwise wouldn't allow it.
  There may be other uses for this, but it is not intended as a general-purpose array,
  as it does not support deleting or altering items. Generally, you will call `Add' as
  many times as necessary to add items to the array, and then use the `Array' or
  `ArrayCopy' properties to retrieve the array built.
)~ */

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
  STDMETHODIMP get_ArrayCopy(/*[out,retval]*/ VARIANT *pvOut);
  STDMETHODIMP Add(/*[in]*/ VARIANT vEl);
  
protected:
  AVAR m_Var;
  UA4  m_Max;
};

OBJECT_ENTRY_AUTO(__uuidof(VarArray), CVarArray)
