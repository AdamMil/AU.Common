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

// UpFile.cpp : Implementation of CUpFile

#include "stdafx.h"
#include "UpFile.h"
#include <cstdio>

// CUpFile

/* ~(MODULES::UPLOAD, p'UpFile::FormName
  <PRE>[propget] HRESULT FormName([out,retval] BSTR *psName);</PRE>
  The FormName property retrieves the form field name of the file.
)~ */
STDMETHODIMP CUpFile::get_FormName(BSTR *psName)
{ if(!psName) return E_POINTER;
  if(!m_Item) return E_UNEXPECTED;
  *psName = m_Item->Name.ToBSTR();
  return S_OK;
} /* get_FormName */

/* ~(MODULES::UPLOAD, p'UpFile::Filename
  <PRE>[propget] HRESULT Filename([out,retval] BSTR *psName);</PRE>
  The Filename property retrieves the file name of the file.
)~ */
STDMETHODIMP CUpFile::get_Filename(BSTR *psName)
{ if(!psName) return E_POINTER;
  if(!m_Item) return E_UNEXPECTED;
  *psName = m_Item->FileName.ToBSTR();
  return S_OK;
} /* get_Filename */

/* ~(MODULES::UPLOAD, p'UpFile::Length
  <PRE>[propget] HRESULT Length([out,retval] long *pnLen);</PRE>
  The Length property retrieves the length of the file, in bytes.
)~ */
STDMETHODIMP CUpFile::get_Length(long *pnLen)
{ if(!pnLen)  return E_POINTER;
  if(!m_Item) return E_UNEXPECTED;
  *pnLen = (long)m_Item->Len;
  return S_OK;
} /* get_Length */

/* ~(MODULES::UPLOAD, p'UpFile::MimeType
  <PRE>[propget] HRESULT MimeType([out,retval] BSTR *psType);</PRE>
  The MimeType property retrieves the mime type of the file, as reported by
  the sending agent.
)~ */
STDMETHODIMP CUpFile::get_MimeType(BSTR *psType)
{ if(!psType) return E_POINTER;
  if(!m_Item) return E_UNEXPECTED;
  *psType = m_Item->ContentType.ToBSTR();
  return S_OK;
} /* get_MimeType */

/* ~(MODULES::UPLOAD, p'UpFile::Data
  <PRE>[propget] HRESULT Data([out,retval] VARIANT *pvData);</PRE>
  The Data property retrieves the data associated with the file as a binary
  array (VT_UI1|VT_ARRAY).
)~ */
STDMETHODIMP CUpFile::get_Data(VARIANT *pvData)
{ if(!pvData) return E_POINTER;
  if(!m_Item) return E_UNEXPECTED;
  SAFEARRAYBOUND bound = { m_Item->Len, 0 };
  SAFEARRAY *sa        = SafeArrayCreate(VT_UI1, 1, &bound);
  if(!sa) return E_FAIL;
  pvData->vt     = VT_ARRAY|VT_UI1;
  pvData->parray = sa;
  std::memcpy(sa->pvData, m_Item->Data, m_Item->Len);
  return S_OK;
} /* get_Data */

/* ~(MODULES::UPLOAD, f'UpFile::Save
  <PRE>HRESULT Data([in] BSTR sPath, [in,defaultvalue(0)] VARIANT_BOOL bCanClobber);</PRE>
  The Save method saves the file to the path specified. The ASP process must have write
  access to that path in order for the call to succeed. You can use `Utility::LogonUser'
  to alter your permissions, if necessary. If 'bCanClobber' is false (the default), an
  error will occur if the file already exists. If 'bCanClobber' is true, the file will
  be overwritten if it already exists.
)~ */
STDMETHODIMP CUpFile::Save(BSTR sPath, VARIANT_BOOL bCanClobber)
{ if(!m_Item) return E_UNEXPECTED;
  HANDLE hFile = CreateFileW(sPath, GENERIC_WRITE, 0, NULL,
                             bCanClobber ? CREATE_ALWAYS : CREATE_NEW,
                             FILE_ATTRIBUTE_NORMAL|FILE_FLAG_SEQUENTIAL_SCAN, NULL);
  if(hFile == INVALID_HANDLE_VALUE) return E_FAIL;
  
  DWORD written;
  WriteFile(hFile, m_Item->Data, m_Item->Len, &written, NULL);
  CloseHandle(hFile);
  return written==m_Item->Len ? S_OK : E_FAIL;
} /* Save */

void CUpFile::Load(CUpload *par, const CUpload::Item *item)
{ m_Parent = par, m_Item = item;
} /* Load */
