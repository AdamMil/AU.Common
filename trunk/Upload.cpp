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

// Upload.cpp : Implementation of CUpload

#include "stdafx.h"
#include "Upload.h"


// CUpload

STDMETHODIMP CUpload::get_Form(BSTR sItem, VARIANT *pvOut)
{ return E_NOTIMPL;
} /* get_Form */

STDMETHODIMP CUpload::get_File(BSTR sItem, IUpFile **ppFile)
{ return E_NOTIMPL;
} /* get_File */

STDMETHODIMP CUpload::get_NumFiles(long *pnFiles)
{ return E_NOTIMPL;
} /* get_NumFiles */

STDMETHODIMP CUpload::GetFile(long nIndex, IUpFile **ppFile)
{ return E_NOTIMPL;
} /* GetFile */

STDMETHODIMP CUpload::OnStartPage(IDispatch *pDisp)
{ return E_NOTIMPL;
} /* OnStartPage */
