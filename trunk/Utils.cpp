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

#include "stdafx.h"
#include "Utils.h"
#include "Config.h"

/*** statics ***/
AComPtr<IConfig> s_Config;
HINSTANCE        s_DllHInst;

/*** rs fields ***/
HRESULT g_GetField(ADORecordset *rs, UA4 nField, VARIANT *out)
{ VARIANT var;
  var.vt=VT_I4, var.intVal=(int)nField;
  return g_GetField(rs, var, out);
} /* g_GetField(ADORecordset *, UA4, VARIANT *) */

HRESULT g_GetField(ADORecordset *rs, VARIANT vField, VARIANT *out)
{ assert(out != NULL);
  if(out==NULL) return E_POINTER;
  AComPtr<ADOFields> fields;
  AComPtr<ADOField>  field;
  HRESULT hRet;
  CHKRET(rs->get_Fields(&fields));
  CHKRET(fields->get_Item(vField, &field));
  CHKRET(field->get_Value(out));
  return hRet;
} /* g_GetField(ADORecordset *, VARIANT, VARIANT *) */

/*** config ***/
HRESULT g_Config(BSTR sKey, BSTR sType, BSTR sSection, AVAR &out)
{ if(!s_Config)
  { HRESULT hRet;
    CHKRET(g_LoadConfig());
  }
  out.Clear();
  return s_Config->get_Item(sKey, sType, sSection, &out);
} /* g_Config */

void g_InitConfig(HINSTANCE hInst)
{ s_DllHInst = hInst;
} /* g_InitConfig */

HRESULT g_LoadConfig()
{ HRESULT hRet;
  char    buf[MAX_PATH+64], *p;
  if(!GetModuleFileName(s_DllHInst, buf, MAX_PATH)) return E_FAIL;
  p=strrchr(buf, L'\\');
  if(!p++) return E_FAIL;
  strcpy(p, "AU.Common.xml");
  if(!s_Config) CHKRET(CREATE(Config, IConfig, s_Config));
  IFF(s_Config->OpenFile(ASTR::FromCStr(buf)))
    hRet=s_Config->OpenXML(ASTR(L"<config/>"));
  return hRet;
} /* g_LoadConfig */

void g_UnloadConfig()
{ // have to wrap this because if the xml dll has already been unloaded,
  // there will be an access violation. the approach is not good, but
  // I don't believe there's a way to control the order of COM server
  // loading/unloading or to detect the application end before
  // CoUnitialize is called
  try { s_Config.Release(); } catch(...) { s_Config.p=NULL; }
} /* g_UnloadConfig */