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

/* ~(MODULES::UTIL, f'Utility::StringToBin
  <PRE>HRESULT StringToBin([in] BSTR sStr, [in,defaultvalue(-1)] VARIANT_BOOL b8bit, [out,retval] VARIANT *pvOut);</PRE>
  The StringToBin method converts a string ('sStr') to a binary array (VT_ARRAY|VT_UI1).
  If 'b8bit' is true (the default), the string is assumed to be comprised of 8-bit
  characters and the high byte of each 16-bit character is ignored. This results
  in a binary array that is one-half the size. However, if you use international
  characters or otherwise store binary data in the upper byte of the character, it
  will be lost. The process can be reversed by `BinToString' with the same 'b8bit'
  value. If the string is empty, NULL (VT_NULL, Nothing) will be returned.
)~ */
STDMETHODIMP CUtility::StringToBin(BSTR sStr, VARIANT_BOOL b8bit, VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  HRESULT hRet;
  AVAR    var;
  CHKRET(g_StringToBin(sStr, b8bit != 0, var));
  *pvOut = var.Detach();
  return S_OK;
} /* StringToBin */

/* ~(MODULES::UTIL, f'Utility::BinToString
  <PRE>HRESULT BinToString([in] VARIANT vBin, [in,defaultvalue(-1)] VARIANT_BOOL b8bit, [out,retval] BSTR *psRet);</PRE>
  The BinToString method converts a binary array (VT_ARRAY of 1-byte elements) to
  a string. If 'b8bit' is true (the default), each byte is expanded into a single
  character in the output string. This will effectively convert ASCII/ANSI 8-bit
  text into a BSTR (although without performing any character set conversion). If
  'b8bit' is false, the binary data is copied directly into the BSTR. This may
  create a BSTR with a non-integral number of characters, if the binary data had
  an odd number of bytes. Unless the binary data is known to be 16-bit character
  data, you should be careful about manipulating the returned string, or else you
  might corrupt the data. If NULL (VT_NULL, Nothing) is passed as vBin, an empty
  string will be returned.
)~ */
STDMETHODIMP CUtility::BinToString(VARIANT vBin, VARIANT_BOOL b8bit, BSTR *psRet)
{ if(!psRet) return E_POINTER;
  HRESULT hRet;
  ASTR    str;
  CHKRET(g_BinToString(vBin, b8bit != 0, str));
  *psRet = str.Detach();
  return hRet;
} /* BinToString */

/* ~(MODULES::UTIL, f'Utility::HexToBin
  <PRE>HRESULT HexToBin([in] BSTR sHex, [out,retval] VARIANT *pvOut);</PRE>
  The HexToBin method converts a string of hex digits into a binary array
  (VT_ARRAY|VT_UI1). If the string is not valid hex, or contains an odd number
  of characters, E_INVALIDARG will be returned. If the string is empty, NULL
  (VT_NULL, Nothing) will be returned.
)~ */
STDMETHODIMP CUtility::HexToBin(BSTR sHex, VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  HRESULT hRet;
  AVAR    var;
  CHKRET(g_HexToBin(sHex, var));
  *pvOut = var.Detach();
  return hRet;
} /* HexToBin */

/* ~(MODULES::UTIL, f'Utility::BinToHex
  <PRE>HRESULT BinToHex([in] VARIANT vBin, [in,defaultvalue(-1)] VARIANT_BOOL b8bit, [out,retval] BSTR *psHex);</PRE>
  The BinToHex method converts a binary array (VT_ARRAY of 1-byte elements) or
  a string to hex. 'b8bit' only has a useful effect when a string is passed in 'vBin'.
  If a string is passed and 'b8bit' is true (the default), the upper 8-bits
  of each character will be ignored. If 'b8bit' is false and a string is
  passed, all the bytes in the string will be converted to
  hex. The result is a hex string that is twice as long but doesn't discard anything.
  See `StringToBin' for more information. It is currently considered an error to
  pass false for 'b8bit' when 'vBin' is a binary array. If NULL (VT_NULL, Nothing)
  is passed as vBin, an empty string will be returned.
)~ */
STDMETHODIMP CUtility::BinToHex(VARIANT vBin, VARIANT_BOOL b8bit, BSTR *psHex)
{ if(!psHex) return E_POINTER;
  HRESULT hRet;
  ASTR    str;
  CHKRET(g_BinToHex(vBin, b8bit != 0, str));
  *psHex = str.Detach();
  return hRet;
} /* BinToHex */

/* ~(MODULES::UTIL, f'Utility::EncodeB64
  <PRE>HRESULT EncodeB64([in] VARIANT vIn, [in,defaultvalue(-1)] VARIANT_BOOL b8bit, [out,retval] BSTR *psB64);</PRE>
  The EncodeB64 method converts a binary array (VT_ARRAY of 1-byte elements) or
  a string to base64. 'b8bit' only has a useful effect when a string is passed in
  'vBin'. If a string is passed and 'b8bit' is true (the default), the upper 8-bits
  of each character will be ignored. If 'b8bit' is false and a string is
  passed, all the bytes in the string will be converted to base64.
  See `StringToBin' for more information. It is currently considered an error to
  pass false for 'b8bit' when 'vBin' is a binary array. If NULL (VT_NULL, Nothing)
  is passed as vBin, an empty string will be returned.
)~ */
STDMETHODIMP CUtility::EncodeB64(VARIANT vIn, VARIANT_BOOL b8bit, BSTR *psB64)
{ if(!psB64) return E_POINTER;
  HRESULT hRet;
  ASTR    str;
  CHKRET(g_EncodeB64(vIn, b8bit != 0, str));
  *psB64 = str.Detach();
  return hRet;
} /* EncodeB64 */

/* ~(MODULES::UTIL, f'Utility::DecodeB64
  <PRE>HRESULT DecodeB64([in] VARIANT vIn, [in,defaultvalue(-1)] VARIANT_BOOL b8bit, [out,retval] VARIANT *pvOut);</PRE>
  The DecodeB64 method converts a binary array (VT_ARRAY of 1-byte elements) or
  a string of base64 bytes. 'b8bit' only has a useful effect when a string is passed in
  'vBin'. If a string is passed and 'b8bit' is true (the default), only the lower 8-bits
  of each character will be considered to be base64. oTHERWISE, if 'b8bit' is false and
  a string is passed, all the bytes in the string will be considered to be
  base64. See `StringToBin' for more information. It is currently considered an error to
  pass false for 'b8bit' when 'vBin' is a binary array. If NULL (VT_NULL, Nothing) or an
  empty string is passed as vBin, NULL will be returned.
)~ */
STDMETHODIMP CUtility::DecodeB64(VARIANT vIn, VARIANT_BOOL b8bit, VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  HRESULT hRet;
  AVAR    var;
  CHKRET(g_DecodeB64(vIn, b8bit != 0, var));
  *pvOut = var.Detach();
  return hRet;
} /* DecodeB64 */

/* ~(MODULES::UTIL, f'Utility::HashSHA1
  <PRE>HRESULT HashSHA1([in] VARIANT vIn, [out,retval] VARIANT *pvOut);</PRE>
  The HashSHA1 method calculates a SHA1 hash for the data passed in vIn. vIn
  can be NULL, a string, or a binary array (VT_ARRAY of one-byte elements).
  The hash returned is a binary array (VT_ARRAY|VT_UI1) of 20 bytes
  (as SHA1 is a 160-bit hash). If a string is passed, all bytes in the string
  are considered.
)~ */
STDMETHODIMP CUtility::HashSHA1(VARIANT vIn, VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  HRESULT hRet;
  AVAR    var;
  CHKRET(g_HashSHA1(vIn, var));
  *pvOut = var.Detach();
  return hRet;
} /* HashSHA1 */

/* ~(MODULES::UTIL, f'Utility::CheckSHA1
  <PRE>HRESULT CheckSHA1([in] VARIANT vIn, [in] VARIANT vHash, [out,retval] VARIANT_BOOL *pbOK);</PRE>
  The CheckSHA1 method works by calling `HashSHA1' on 'vIn' to get a hash, and then
  comparing the hash with 'vHash'. If the hashes match, the function returns true.
  Otherwise, it returns false.
)~ */
STDMETHODIMP CUtility::CheckSHA1(VARIANT vIn, VARIANT vHash, VARIANT_BOOL *pbOK)
{ if(!pbOK) return E_POINTER;
  HRESULT hRet;
  bool    ok;
  CHKRET(g_CheckSHA1(vIn, vHash, ok));
  *pbOK = ok ? VBTRUE : VBFALSE;
  return hRet;
} /* CheckSHA1 */

/* ~(MODULES::UTIL, f'Utility::VarType
  <PRE>HRESULT VarType([in] VARIANT vIn, [out,retval] short *pnType);</PRE>
  The VarType method simply returns the .vt member of the variant passed.
  This allows the exact variant type to be determined.
)~ */
STDMETHODIMP CUtility::VarType(VARIANT vIn, short *pnType)
{ if(!pRet) return E_POINTER;
  *pRet = vIn.vt;
  return S_OK;
} /* VarType */

/* ~(MODULES::UTIL, f'Utility::GetTickCount
  <PRE>HRESULT GetTickCount([[out,retval] long *pnMS);</PRE>
  The GetTickCount method calls the Win32 API function GetTickCount and
  returns the value obtained. This is the number of milliseconds since the
  system has been rebooted. It wraps around after approximately 47 days.
)~ */
STDMETHODIMP CUtility::GetTickCount(long *pnMS)
{ if(!pnMS) return E_POINTER;
  *pnMS = (long)::GetTickCount();
  return S_OK;
} /* GetTickCount */

/* ~(MODULES::UTIL, f'Utility::LogonUser
  <PRE>HRESULT LogonUser([in] BSTR sUser, [in] BSTR sPass, [in,defaultvalue("")] BSTR sDomain);</PRE>
  The LogonUser method allows the caller to impersonate a user in order to access
  resources not normally accessible. This is useful, for instance, from ASP in
  order to quickly get administrative priviledges. 'sUser' and 'sPass' are the
  login for a system user. If the user is on another domain, 'sDomain' can be
  used to specify the domain in which the user resides. If used from ASP, 
  it is important to call `RevertLogin' or to destroy the Utility object before
  control is returned to ASP (the end of the page). It's best to call `RevertLogin'
  as soon as you have the information you need.
)~ */
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

/* ~(MODULES::UTIL, f'Utility::RevertLogin
  <PRE>HRESULT RevertLogin();</PRE>
  The RevertLogin method releases the login token acquired by `LogonUser'. This is
  called when the Utility object is destroyed, if necessary, but should be called
  expicitly so that you can be sure that things aren't accessed/called with improper
  security credentials.
)~ */
STDMETHODIMP CUtility::RevertLogin()
{ if(m_hToken != NULL)
  { RevertToSelf();
    CloseHandle(m_hToken);
    m_hToken = NULL;
    return S_OK;
  }
  return S_FALSE;
} /* RevertLogin */
