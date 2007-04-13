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

/* ~(!CATEGORIES
  INTRODUCTION=Introduction
  MODULES=AU.Common Modules
  MODULES::CONFIG=Config Object
  MODULES::DB=DB Object
  MODULES::DBMAN=DBMan Object
  MODULES::UTIL=Utility Object
  MODULES::UPLOAD=Upload/UpFile Objects
  MODULES::VARARR=VarArray Object
  MODULES::INC=Server-side Javascript Libraries
  MODULES::INC::BROWSER=Browser detection
  MODULES::INC::VALIDATE=Validation package
  MODULES::JS=Client-side Javascript Libraries
  MODULES::JS::HOOKS=Event chaining
  MODULES::JS::VALIDATE=Client-side Validation Support
)~ */

// Common.cpp : Implementation of DLL Exports.

#include "stdafx.h"
#include "resource.h"
#include "Common.h"

class CCommonModule : public CAtlDllModuleT< CCommonModule >
{
public :
  DECLARE_LIBID(LIBID_AU_CommonLib)
  DECLARE_REGISTRY_APPID_RESOURCEID(IDR_COMMON, "{4BFCDF23-2787-401A-BF17-25B14E84E38F}")
};

CCommonModule _AtlModule;


// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInst, DWORD dwReason, LPVOID lpReserved)
{ BOOL ret;
  static int nInit;
  if(dwReason==DLL_PROCESS_DETACH && !--nInit) g_UnloadConfig();
  ret = _AtlModule.DllMain(dwReason, lpReserved);
  if(dwReason==DLL_PROCESS_ATTACH && !nInit++)
  { 
  #ifndef NDEBUG
    _CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_CHECK_ALWAYS_DF |
                    _CRTDBG_LEAK_CHECK_DF);
  #endif
    g_vMissing.vt=VT_ERROR, g_vMissing.scode=DISP_E_PARAMNOTFOUND;
    g_InitConfig(hInst);
  }
  return ret;
}


// Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow(void)
{ return _AtlModule.DllCanUnloadNow();
}


// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{ return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}


// DllRegisterServer - Adds entries to the system registry
STDAPI DllRegisterServer(void)
{
    // registers object, typelib and all interfaces in typelib
    HRESULT hr = _AtlModule.DllRegisterServer();
  return hr;
}


// DllUnregisterServer - Removes entries from the system registry
STDAPI DllUnregisterServer(void)
{
  HRESULT hr = _AtlModule.DllUnregisterServer();
  return hr;
}
