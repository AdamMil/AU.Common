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

// Utility.cpp : Implementation of CUtility

#include "stdafx.h"
#include "Utility.h"

// CUtility
CUtility::CUtility()
{ m_hToken = NULL;
} /* CUtility */

CUtility::~CUtility()
{ if(m_hToken != NULL)
  { RevertToSelf();
    CloseHandle(m_hToken);
  }
} /* ~CUtility */

STDMETHODIMP CUtility::StringToBin(BSTR sStr, VARIANT_BOOL b8bit, VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  HRESULT hRet;
  AVAR    var;
  CHKRET(g_StringToBin(sStr, b8bit != 0, var));
  *pvOut = var.Detach();
  return S_OK;
} /* StringToBin */

STDMETHODIMP CUtility::BinToString(VARIANT vBin, VARIANT_BOOL b8bit, BSTR *psRet)
{ if(!psRet) return E_POINTER;
  HRESULT hRet;
  ASTR    str;
  CHKRET(g_BinToString(vBin, b8bit != 0, str));
  *psRet = str.Detach();
  return hRet;
} /* BinToString */

STDMETHODIMP CUtility::HexToBin(BSTR sHex, VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  HRESULT hRet;
  AVAR    var;
  CHKRET(g_HexToBin(sHex, var));
  *pvOut = var.Detach();
  return hRet;
} /* HexToBin */

STDMETHODIMP CUtility::BinToHex(VARIANT vBin, VARIANT_BOOL b8bit, BSTR *psHex)
{ if(!psHex) return E_POINTER;
  HRESULT hRet;
  ASTR    str;
  CHKRET(g_BinToHex(vBin, b8bit != 0, str));
  *psHex = str.Detach();
  return hRet;
} /* BinToHex */

STDMETHODIMP CUtility::EncodeB64(VARIANT vIn, VARIANT_BOOL b8bit, VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  HRESULT hRet;
  AVAR    var;
  CHKRET(g_EncodeB64(vIn, b8bit != 0, var));
  *pvOut = var.Detach();
  return hRet;
} /* EncodeB64 */

STDMETHODIMP CUtility::DecodeB64(VARIANT vIn, VARIANT_BOOL b8bit, VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  HRESULT hRet;
  AVAR    var;
  CHKRET(g_DecodeB64(vIn, b8bit != 0, var));
  *pvOut = var.Detach();
  return hRet;
} /* DecodeB64 */

// 512mb limit
STDMETHODIMP CUtility::HashSHA1(VARIANT vIn, VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  HRESULT hRet;
  AVAR    var;
  CHKRET(g_HashSHA1(vIn, var));
  *pvOut = var.Detach();
  return hRet;
} /* HashSHA1 */

STDMETHODIMP CUtility::CheckSHA1(VARIANT vIn, VARIANT vHash, VARIANT_BOOL *pbOK)
{ if(!pbOK) return E_POINTER;
  HRESULT hRet;
  bool    ok;
  CHKRET(g_CheckSHA1(vIn, vHash, ok));
  *pbOK = ok ? VBTRUE : VBFALSE;
  return hRet;
} /* CheckSHA1 */

STDMETHODIMP CUtility::VarType(VARIANT vIn, short *pRet)
{ if(!pRet) return E_POINTER;
  *pRet = vIn.vt;
  return S_OK;
} /* VarType */

STDMETHODIMP CUtility::GetTickCount(long *pnMS)
{ if(!pnMS) return E_POINTER;
  *pnMS = (long)::GetTickCount();
  return S_OK;
} /* GetTickCount */

STDMETHODIMP CUtility::LogonUser(BSTR sUser, BSTR sPass, BSTR sDomain)
{ if(m_hToken != NULL) RevertLogin();
  else RevertToSelf();
  
  HRESULT hRet = E_ACCESSDENIED;
  if(LogonUserW(sUser, sDomain&&sDomain[0] ? sDomain : NULL, sPass,
                LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, &m_hToken))
  { if(!ImpersonateLoggedOnUser(m_hToken))
    { CloseHandle(m_hToken);
      m_hToken = NULL;
    }
    else hRet=S_OK;
  }
  return hRet;
} /* LogonUser */

STDMETHODIMP CUtility::RevertLogin()
{ if(m_hToken != NULL)
  { RevertToSelf();
    CloseHandle(m_hToken);
    m_hToken = NULL;
    return S_OK;
  }
  return S_FALSE;
} /* RevertLogin */
