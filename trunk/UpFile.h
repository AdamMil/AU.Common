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

// UpFile.h : Declaration of the CUpFile

#pragma once
#include "resource.h"       // main symbols

#include "Common.h"


// CUpFile

class ATL_NO_VTABLE CUpFile : 
  public CComObjectRootEx<CComSingleThreadModel>,
  public CComCoClass<CUpFile, &CLSID_UpFile>,
  public IDispatchImpl<IUpFile, &IID_IUpFile, &LIBID_AU_CommonLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
  DECLARE_REGISTRY_RESOURCEID(IDR_UPFILE)
  DECLARE_NOT_AGGREGATABLE(CUpFile)
  BEGIN_COM_MAP(CUpFile)
    COM_INTERFACE_ENTRY(IUpFile)
    COM_INTERFACE_ENTRY(IDispatch)
  END_COM_MAP()
  DECLARE_PROTECT_FINAL_CONSTRUCT()

// IUpFile
public:
  STDMETHODIMP get_FormName(/*[out,retval]*/ BSTR *psName);
  STDMETHODIMP get_Filename(/*[out,retval]*/ BSTR *psName);
  STDMETHODIMP get_Length(/*[out,retval]*/ long *pnLen);
  STDMETHODIMP get_MimeType(/*[out,retval]*/ BSTR *psType);
  STDMETHODIMP get_Attribute(/*[in]*/ BSTR sAttr, /*[out,retval]*/ BSTR *psVal);
  STDMETHODIMP get_Data(/*[out,retval]*/ VARIANT *pvData);
  STDMETHODIMP Save(/*[in]*/ BSTR sPath, /*[in,defaultvalue(0)]*/ VARIANT_BOOL bCanClobber);
};

OBJECT_ENTRY_AUTO(__uuidof(UpFile), CUpFile)
