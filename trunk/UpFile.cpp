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

// UpFile.cpp : Implementation of CUpFile

#include "stdafx.h"
#include "UpFile.h"


// CUpFile

STDMETHODIMP CUpFile::get_FormName(BSTR *psName)
{
} /* get_FormName */

STDMETHODIMP CUpFile::get_Filename(BSTR *psName)
{
} /* get_Filename */

STDMETHODIMP CUpFile::get_Length(long *pnLen)
{
} /* get_Length */

STDMETHODIMP CUpFile::get_MimeType(BSTR *psType)
{
} /* get_MimeType */

STDMETHODIMP CUpFile::get_Attribute(BSTR sAttr, BSTR *psVal)
{
} /* get_Attribute */

STDMETHODIMP CUpFile::get_Data(VARIANT *pvData)
{
} /* get_Data */

STDMETHODIMP CUpFile::Save(BSTR sPath, VARIANT_BOOL bCanClobber)
{
} /* Save */
