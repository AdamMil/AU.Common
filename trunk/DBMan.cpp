/* 
   This file is part of the AU.Common library, a set of ActiveX
   controls and C++ classes used to aid in COM and Web development.
   Copyright (C) 2002 COM and Adam Milazzo

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

// DBMan.cpp : Implementation of CDBMan

#include "stdafx.h"
#include "DBMan.h"

// CDBMan

STDMETHODIMP CDBMan::get_MaxSharing(long *pnShares)
{ if(!pnShares) return E_POINTER;
  *pnShares = (long)m_MaxShares;
  return S_OK;
} /* get_MaxSharing */

STDMETHODIMP CDBMan::put_MaxSharing(long nShares);
{ if(nShares<0) return E_INVALIDARG;
  m_MaxShares = (UA4)nShares;
  return S_OK;
} /* put_MaxSharing */

STDMETHODIMP CDBMan::CreateDB(BSTR sKey, BSTR sSect, IDB **ppDB);
{ assert(sKey && sSect);
  AVAR var;
  CHKRET(g_Config(sKey[0] ? sKey : ASTR(L"DB/Default"), ASTR(L"string"), sSect, var));
  return CreateDBRaw(var.ToBSTR(), ppDB);
} /* CreateDB */

STDMETHODIMP CDBMan::CreateDBRaw(BSTR sConnStr, IDB **ppDB);
{ if(!sConnStr || !sConnStr[0]) return E_INVALIDARG;
  
} /* CreateDBRaw */
