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

#ifndef AU_COMMON_UTILS_H
#define AU_COMMON_UTILS_H

HRESULT g_GetField(ADORecordset *, UA4 nField, VARIANT *out);
HRESULT g_GetField(ADORecordset *, VARIANT vField, VARIANT *out);

HRESULT g_Config(BSTR sKey, BSTR sType, BSTR sSection, AVAR &out);
void    g_InitConfig(HINSTANCE hInst);
HRESULT g_LoadConfig();
void    g_UnloadConfig();

HRESULT g_StringToBin(BSTR sStr, bool b8bit /*=true*/, AVAR &out);
HRESULT g_BinToString(VARIANT &vBin, bool b8bit /*=true*/, ASTR &out);
HRESULT g_HexToBin(BSTR sHex, AVAR &out);
HRESULT g_BinToHex(VARIANT &vBin, bool b8bit /*=true*/, ASTR &out);
HRESULT g_EncodeB64(VARIANT &vIn, bool b8bit /*=true*/, ASTR &out);
HRESULT g_DecodeB64(VARIANT &vIn, bool b8bit /*=true*/, AVAR &out);
HRESULT g_HashSHA1(VARIANT &vIn, AVAR &out);
HRESULT g_CheckSHA1(VARIANT &vIn, VARIANT &vHash, bool &ok);

#endif
