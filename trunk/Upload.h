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

#include "Common.h"


/* ~(MODULES::UPLOAD, c'Upload
  The Upload object (ProgID AU.Common.Upload) provides a simple way to access
  files uploaded to an ASP page. The Upload object should be created on the
  page with a Server.CreateObject() call. Since the object uses the raw request
  data, the Request.Form collection cannot be accessed if the Upload object has
  been created. ASP will throw an exception. Similarly, if the Request.Form
  collection has been accessed, the object will throw an exception upon creation.
  However, you can use the `Form', `NumFields', and `FieldKey' properties to
  retrieve form data as well as iterate through the form. Files can be accessed
  by ordinal or by form field name. Files are represented by `UpFile' objects.
  In order for file uploading to work, you need to use the multipart/form-data
  encoding type for your form. For example:
  <PRE><form method="post" enctype="multipart/form-data">...</form></PRE>
  The Upload object uses the "apartment" threading model. LIMITATIONS: This
  object does not support multiple fields/files with the same name (only one
  will be seen) or filenames with embedded quotes.
  The Upload object uses the apartment threading model.
)~ */

class ATL_NO_VTABLE CUpload : 
  public CComObjectRootEx<CComSingleThreadModel>,
  public CComCoClass<CUpload, &CLSID_Upload>,
  public IDispatchImpl<IUpload, &IID_IUpload, &LIBID_AU_CommonLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
  DECLARE_REGISTRY_RESOURCEID(IDR_UPLOAD)
  DECLARE_NOT_AGGREGATABLE(CUpload)
  BEGIN_COM_MAP(CUpload)
    COM_INTERFACE_ENTRY(IUpload)
    COM_INTERFACE_ENTRY(IDispatch)
  END_COM_MAP()
  DECLARE_PROTECT_FINAL_CONSTRUCT()

// IUpload
public:
  STDMETHODIMP get_Form(/*[in]*/ BSTR sKey, /*[out,retval]*/ VARIANT *pvOut);
  STDMETHODIMP get_File(/*[in]*/ VARIANT vIndex, /*[out,retval]*/ IUpFile **ppFile);
  STDMETHODIMP get_NumFiles(/*[out,retval]*/ long *pnFiles);
  STDMETHODIMP get_NumFields(/*[out,retval]*/ long *pnFields);
  STDMETHODIMP get_FieldKey(/*[in]*/ long nIndex, /*[out,retval]*/ BSTR *psField);
  STDMETHODIMP OnStartPage(/*[in]*/ IDispatch *pDisp);
  
protected:
  UA4 URLDecode(char *start, char *end);

  struct Item
  { ASTR  Name, ContentType, FileName;
    U1   *Data;
    UA4   Len;
  };
  
  typedef std::map<ASTR, Item> Map;
  typedef std::pair<ASTR, Item> MapPair;

  AVAR m_vData;
  Map  m_mForm, m_mFiles;
  
  friend class CUpFile;
};

OBJECT_ENTRY_AUTO(__uuidof(Upload), CUpload)
