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

// Upload.h : Declaration of the CUpload

#pragma once
#include "resource.h"       // main symbols

#include "AU.Common.h"


// CUpload

class ATL_NO_VTABLE CUpload : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CUpload, &CLSID_Upload>,
	public IDispatchImpl<IUpload, &IID_IUpload, &LIBID_AU_CommonLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CUpload()

  DECLARE_REGISTRY_RESOURCEID(IDR_UPLOAD)
  DECLARE_NOT_AGGREGATABLE(CUpload)
  BEGIN_COM_MAP(CUpload)
	  COM_INTERFACE_ENTRY(IUpload)
	  COM_INTERFACE_ENTRY(IDispatch)
  END_COM_MAP()
  DECLARE_PROTECT_FINAL_CONSTRUCT()

// IUpload
public:
  STDMETHODIMP get_Form(/*[in]*/ BSTR sItem, /*[out,retval]*/ VARIANT *pvOut);
  STDMETHODIMP get_File(/*[in]*/ BSTR sItem, /*[out,retval]*/ IUpFile **ppFile);
  STDMETHODIMP get_NumFiles(/*[out,retval]*/ long *pnFiles);
  STDMETHODIMP GetFile(/*[in]*/ long nIndex, /*[out,retval]*/ IUpFile **ppFile);
  STDMETHODIMP OnStartPage(/*[in]*/ IDispatch *pDisp);
  
protected:
  
};

OBJECT_ENTRY_AUTO(__uuidof(Upload), CUpload)
