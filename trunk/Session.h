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

// Session.h : Declaration of the CAUSession

#pragma once
#include "resource.h"       // main symbols

#include "Common.h"
#include "BlockAlloc.h"

class CSessionMan;

// NOT SAFE TO USE SECTION NAME WITH '%'
// TEXT DATA CURRENTLY LIMITED TO 8k

// CAUSession
class ATL_NO_VTABLE CAUSession : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CAUSession, &CLSID_AUSession>,
	public IDispatchImpl<IAUSession, &IID_IAUSession, &LIBID_AU_CommonLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
  CAUSession();
 ~CAUSession();

  DECLARE_REGISTRY_RESOURCEID(IDR_SESSION)
  DECLARE_NOT_AGGREGATABLE(CAUSession)
  BEGIN_COM_MAP(CAUSession)
	  COM_INTERFACE_ENTRY(IAUSession)
	  COM_INTERFACE_ENTRY(IDispatch)
	  COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
  END_COM_MAP()
	DECLARE_PROTECT_FINAL_CONSTRUCT()
	DECLARE_GET_CONTROLLING_UNKNOWN()

	HRESULT FinalConstruct() { return CoCreateFreeThreadedMarshaler(GetControllingUnknown(), &m_pUnkMarshaler.p); }
	void FinalRelease();

	CComPtr<IUnknown> m_pUnkMarshaler;
public:
  enum Flags { SF_PERSIST=1 };

  STDMETHODIMP get_Item(/*[in]*/ BSTR sKey, /*[out,retval]*/ VARIANT *pvOut);
  STDMETHODIMP put_Item(/*[in]*/ BSTR sKey, /*[in]*/ VARIANT vData);
  STDMETHODIMP get_SessionID(/*[out,retval]*/ BSTR *psID);
  STDMETHODIMP get_Section(/*[out,retval]*/ BSTR *psSect);
  STDMETHODIMP put_Section(/*[in]*/ BSTR sSect);
  STDMETHODIMP get_Flags(/*[out,retval]*/ long *pnFlags);
  STDMETHODIMP put_Flags(/*[in]*/ long nFlags);
  STDMETHODIMP get_Timeout(/*[out,retval]*/ long *pnSecs);
  STDMETHODIMP put_Timeout(/*[in]*/ long nSecs);
  STDMETHODIMP get_IsManaged(/*[out,retval]*/ VARIANT_BOOL *pbManaged);
  STDMETHODIMP get_UsesDB(/*[out,retval]*/ VARIANT_BOOL *pbDB);
  STDMETHODIMP put_UsesDB(/*[in]*/ VARIANT_BOOL bDB);
  STDMETHODIMP ResetSectEnum();
  STDMETHODIMP ResetKeyEnum(/*[in]*/ BSTR sSect);
  STDMETHODIMP get_NextSection(/*[out,retval]*/ VARIANT *pvSect);
  STDMETHODIMP get_NextKey(/*[out,retval]*/ VARIANT *pvKey);

  STDMETHODIMP Clear(/*[in]*/ BSTR sSect);
  STDMETHODIMP ClearAll();
  STDMETHODIMP Persist(/*[out,retval]*/ BSTR *psID);
  STDMETHODIMP Revert();
  STDMETHODIMP Save();
  STDMETHODIMP Delete();
  STDMETHODIMP LockSession()   { Lock();   return S_OK; }
  STDMETHODIMP UnlockSession() { Unlock(); return S_OK; }
  STDMETHODIMP LoadKeysFromDB();
  
protected:
  enum { DEFTIMEOUT=20*60 };
  struct Item
  { VARIANT Var;
    U4      ID;
    bool    bDirty;
  };
  typedef std::map <ASTR, Item*> KeyMap;
  typedef std::pair<ASTR, Item*> KeyMapPair;
  typedef std::map <ASTR, KeyMap*> SectMap;
  typedef std::pair<ASTR, KeyMap*> SectMapPair;

  HRESULT  InitSession(bool useDB=false);
  Item *   FindItem(const ASTR &key, bool autoInsert=true);
  HRESULT  BecomeDB(bool persist=false);
  void     DeleteSection(SectMap::iterator);
  void     DeleteSection2(SectMap::iterator);
  HRESULT  DeleteFromDB();
  HRESULT  AddDBItem(const ASTR &skey, const ASTR &mkey, Item *);
  HRESULT  AddDBItem(const ASTR &key, Item *);
  HRESULT  ReadDBItem(Item *, VARIANT *);
  HRESULT  WriteDBItem(Item *);
  HRESULT  GetDBItem(ASTR &key, Item **);
  HRESULT  ReadValue(ADORecordset *rs, VARIANT *pvout);
  SectMap::iterator FindSection(const ASTR &key, bool autoInsert=true);
  SectMap::iterator SplitKey(ASTR &key);
    
  SectMap m_Sects;
  SectMap::iterator m_CurSect, m_EnumSect;
  KeyMap ::iterator m_EnumKey;
  KeyMap *m_EnumKeyMap;
  ASTR    m_SessionID;
  AComPtr<CSessionMan> m_Par;
  AComPtr<IDB>         m_DB;
  U4      m_Flags, m_Timeout;
  long    m_DBID;
  bool    m_bNoInsert;

  static ACritSec          s_AllocCS;
  static CBlockAlloc<Item> s_ItemAlloc;
};

OBJECT_ENTRY_AUTO(__uuidof(AUSession), CAUSession)
