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

// Session.cpp : Implementation of CAUSession

#include "stdafx.h"
#include "Session.h"
#include "SessionMan.h"

// CAUSession
CAUSession::CAUSession()
{ m_Flags   = 0;
  m_Timeout = DEFTIMEOUT;
  m_DBID    = 0;
} /* CAUSession */

CAUSession::~CAUSession()
{ ClearAll();
  SectMap::iterator si = m_Sects.begin();
  for(; si != m_Sects.end(); si++) delete (*si).second;
} /* ~CAUSession */

void CAUSession::FinalRelease()
{ Save();
  m_pUnkMarshaler.Release();
} /* FinalRelease */

STDMETHODIMP CAUSession::get_Item(BSTR sKey, VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  AComLock lock(this);
  HRESULT  hRet;
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  ASTR  key(sKey);
  Item *it = FindItem(key.LowerCase(), m_DBID==0);
  if(it==NULL)
  { CHKRET(GetDBItem(key, &it));
    if(hRet==S_FALSE) pvOut->vt=VT_NULL;
    else hRet=VariantCopy(pvOut, &it->Var);
  }
  else hRet=VariantCopy(pvOut, &it->Var);
  return hRet;
} /* get_Item */

STDMETHODIMP CAUSession::put_Item(BSTR sKey, VARIANT vData)
{ AComLock lock(this);
  HRESULT  hRet;
  if(m_DBID) CHKRET(g_DBCheckType(&vData));
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  Item *it = FindItem(ASTR(sKey).LowerCase());
  VariantClear(&it->Var);
  VariantCopy(&it->Var, &vData);
  if(m_DBID && (m_Flags&SF_WRITETHROUGH))
  { hRet = WriteDBItem(it);
    it->bDirty = false;
  }
  else it->bDirty = true;
  return hRet;
} /* put_Item */

STDMETHODIMP CAUSession::get_SessionID(BSTR *psID)
{ if(!psID) return E_POINTER;
  AComLock lock(this);
  HRESULT  hRet;
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  *psID = m_SessionID.ToBSTR();
  return S_OK;
} /* get_SessionID */

STDMETHODIMP CAUSession::get_Section(BSTR *psSect)
{ if(!psSect) return E_POINTER;
  AComLock lock(this);
  HRESULT  hRet;
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  *psSect = (*m_CurSect).first.ToBSTR();
  return S_OK;
} /* get_Section */

STDMETHODIMP CAUSession::put_Section(BSTR sSect)
{ AComLock lock(this);
  HRESULT  hRet;
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  m_CurSect = FindSection(ASTR(sSect).LowerCase());
  return S_OK;
} /* put_Section */

STDMETHODIMP CAUSession::get_Flags(long *pnFlags)
{ if(!pnFlags) return E_POINTER;
  *pnFlags = m_Flags;
  return S_OK;
} /* get_Flags */

STDMETHODIMP CAUSession::put_Flags(long nFlags)
{ AComLock lock(this);
  if((m_Flags & SF_PERSIST) && (nFlags & SF_PERSIST)==0 ||
     (m_Flags & SF_PERSIST)==0 && (nFlags & SF_PERSIST))
    return E_INVALIDARG;
  m_Flags = nFlags;
  return S_OK;
} /* put_Flags */

STDMETHODIMP CAUSession::get_Timeout(long *pnSecs)
{ if(!pnSecs) return E_POINTER;
  *pnSecs = m_Timeout;
  return S_OK;
} /* get_Timeout */

STDMETHODIMP CAUSession::put_Timeout(long nSecs)
{ if(nSecs < 0) return E_INVALIDARG;
  AComLock lock(this);
  m_Timeout = nSecs;
  return S_OK;
} /* put_Timeout */

STDMETHODIMP CAUSession::get_IsManaged(VARIANT_BOOL *pbManaged)
{ if(!pbManaged) return E_POINTER;
  *pbManaged = (m_Par==NULL ? VBFALSE : VBTRUE);
  return S_OK;
} /* get_IsManaged */

STDMETHODIMP CAUSession::get_UsesDB(VARIANT_BOOL *pbDB)
{ if(!pbDB) return E_POINTER;
  HRESULT  hRet;
  AComLock lock(this);
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  *pbDB = (m_DBID==0 ? VBFALSE : VBTRUE);
  return S_OK;
} /* get_UsesDB */

STDMETHODIMP CAUSession::put_UsesDB(VARIANT_BOOL bDB)
{ AComLock lock(this);
  HRESULT  hRet;
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  if(m_Flags & SF_PERSIST) return E_UNEXPECTED;
  if(m_DBID)
  { if(bDB) return S_FALSE;
    CHKRET(DeleteFromDB());
    m_DBID = 0;
  }
  else
  { if(!bDB) return S_FALSE;
    hRet = BecomeDB();
  }
  return hRet;
} /* put_UsesDB */

STDMETHODIMP CAUSession::ResetSectEnum()
{ AComLock lock(this);
  HRESULT  hRet;
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  m_EnumSect = m_Sects.begin();
  return S_OK;
} /* ResetSectEnum */

STDMETHODIMP CAUSession::ResetKeyEnum(BSTR sSect)
{ AComLock lock(this);
  HRESULT  hRet;
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  SectMap::iterator it = FindSection(ASTR(sSect).LowerCase(), false);
  if(it==m_Sects.end()) return E_INVALIDARG;
  m_EnumKeyMap = (*it).second;
  m_EnumKey = m_EnumKeyMap->begin();
  return S_OK;
} /* ResetKeyEnum */

STDMETHODIMP CAUSession::get_NextSection(VARIANT *pvSect)
{ AComLock lock(this);
  HRESULT  hRet;
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  if(m_EnumSect==m_Sects.end())
  { pvSect->vt=VT_NULL;
    return S_FALSE;
  }
  pvSect->vt = VT_BSTR;
  pvSect->bstrVal = (*m_EnumSect).first.ToBSTR();
  m_EnumSect++;
  return S_OK;
} /* get_NextSection */

STDMETHODIMP CAUSession::get_NextKey(VARIANT *pvKey)
{ AComLock lock(this);
  HRESULT  hRet;
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  if(m_EnumKey==m_EnumKeyMap->end())
  { pvKey->vt=VT_NULL;
    return S_FALSE;
  }
  pvKey->vt = VT_BSTR;
  pvKey->bstrVal = (*m_EnumKey).first.ToBSTR();
  m_EnumKey++;
  return S_OK;
} /* get_NextKey */

STDMETHODIMP CAUSession::Clear(BSTR sSect)
{ AComLock lock(this);
  HRESULT  hRet;
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  SectMap::iterator it = FindSection(ASTR(sSect).LowerCase(), false);
  if(it==m_Sects.end()) return E_INVALIDARG;

  bool bResetKey, bResetSect;
  if(!(m_EnumSect!=it)) bResetSect=true;
  if((*it).second==m_EnumKeyMap) bResetKey=true;

  if(!(m_CurSect!=it))
  { DeleteSection2(it);
    m_CurSect = FindSection(ASTR());
  }
  else DeleteSection2(it);

  if(bResetSect) ResetSectEnum();
  if(bResetKey)  ResetKeyEnum((BSTR)ASTR::EmptyBSTR);
  return S_OK;
} /* Clear */

STDMETHODIMP CAUSession::ClearAll()
{ AComLock lock(this);
  HRESULT  hRet;
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  SectMap::iterator it, next = m_Sects.begin();

  if(m_DBID)
  { SAFEARRAY sa, *psa=&sa;
    VARIANT   var;
    g_InitSA(&sa, &var);
    var.vt = VT_I4, var.intVal = m_DBID;
    CHKRET(m_DB->ExecuteONR(ASTR(L"sess_Clear"), L"@session_id", &psa));
    while(next != m_Sects.end()) DeleteSection(next++);
  }
  else while(next != m_Sects.end()) DeleteSection(next++);
  m_CurSect    = m_EnumSect = FindSection(ASTR());
  m_EnumKeyMap = (*m_CurSect).second;
  m_EnumKey    = m_EnumKeyMap->begin();
  return S_OK;
} /* ClearAll */

STDMETHODIMP CAUSession::Persist(BSTR *psID)
{ AComLock lock(this);
  HRESULT  hRet;
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  if(m_Flags&SF_PERSIST) return S_FALSE;
  CHKRET(BecomeDB());
  m_Flags |= SF_PERSIST;
  return hRet;
} /* Persist */

STDMETHODIMP CAUSession::Revert()
{ AComLock lock(this);
  HRESULT  hRet=S_OK;
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  if(!m_DBID) return E_UNEXPECTED;
  
  SectMap::iterator s = m_Sects.begin();
  KeyMap::iterator  k;
  KeyMap           *km;
  Item             *it;
  VARIANT           var;
  
  for(; s != m_Sects.end(); s++)
  { km = (*s).second;
    k  = km->begin();
    for(; k != km->end(); k++)
    { it = (*k).second;
      if(it->bDirty)
      { if(it->ID)
        { CHKRT2(ReadDBItem(it, &var), Done);
          VariantClear(&it->Var);
          it->Var = var;
        }
        else VariantClear(&it->Var);
        it->bDirty = false;
      }
    }
  }
  Done:
  return hRet;
} /* Revert */

STDMETHODIMP CAUSession::Save()
{ AComLock lock(this);
  HRESULT  hRet=S_OK;
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  if(!m_DBID) return E_UNEXPECTED;

  SectMap::iterator s = m_Sects.begin();
  KeyMap::iterator  k;
  KeyMap           *km;
  Item             *it;
  
  for(; s != m_Sects.end(); s++)
  { km = (*s).second;
    k  = km->begin();
    for(; k != km->end(); k++)
    { it = (*k).second;
      if(it->bDirty && FAILED(WriteDBItem(it))) hRet=E_FAIL;
    }
  }
  return hRet;
} /* Save */

STDMETHODIMP CAUSession::Delete()
{ AComLock lock(this);
  HRESULT  hRet;
  // don't need to initialize
  if(m_DBID)
  { CHKRET(DeleteFromDB());
    ClearAll();
    m_DBID = 0;
    hRet   = S_OK;
  }
  else hRet = S_FALSE;
  m_SessionID.Clear();
  m_Flags   = 0;
  m_Timeout = DEFTIMEOUT;
  return hRet;
} /* Delete */

STDMETHODIMP CAUSession::LoadKeysFromDB()
{ AComLock lock(this);
  HRESULT  hRet;
  if(m_SessionID.IsEmpty()) CHKRET(InitSession());
  if(!m_DBID) return E_UNEXPECTED;

  ARecordset   rs;
  BSTR         key;
  SAFEARRAY    sa, *psa=&sa;
  VARIANT      var;
  VARIANT_BOOL eof=1;
  bool         failed=false;

  g_InitSA(&sa, &var);
  var.vt = VT_I4, var.intVal = m_DBID;
  CHKRET(m_DB->ExecuteO(ASTR(L"sess_LoadSession"), L"@session_id", &psa, &rs));

  REWRITE ME
  while(rs->get_EOF(&eof), !eof)
  { VariantClear(&var);
    CHKRT2(g_GetField(rs, 0, &var), Done);
    key =var.bstrVal, var.vt=VT_EMPTY;
    hRet=get_Item(key, &var);
    SysFreeString(key);
    if(FAILED(hRet)) goto Done;
    rs->MoveNext();
  }
  Done:
  return hRet;
} /* LoadKeysFromDB */

HRESULT CAUSession::InitSession(bool useDB)
{ if(m_Flags & SF_PERSIST) useDB=true;
  
  m_CurSect    = m_EnumSect = FindSection(ASTR());
  m_EnumKeyMap = (*m_CurSect).second;
  m_EnumKey    = m_EnumKeyMap->begin();

  if(useDB) BecomeDB();
  else
  { if(m_Par != NULL)
    { // TODO: let parent generate key to guarantee uniqueness among managed sessions
    }
    else
    { WCHAR key[21];
      g_RandStr(key, 20); key[20] = 0;
      m_SessionID = key;
    }
  }
  return S_OK;
} /* InitSession */

CAUSession::Item * CAUSession::FindItem(const ASTR &key, bool autoInsert)
{ ASTR   mkey(key);
  KeyMap  *k = (*(SplitKey(mkey))).second;
  KeyMap::iterator ki = k->find(mkey);
  if(ki == k->end())
  { if(autoInsert)
    { s_AllocCS.Lock();
      Item *it = s_ItemAlloc.New();
      s_AllocCS.Unlock();
      it->Var.vt = VT_EMPTY;
      it->ID     = 0;
      it->bDirty = false;
      ki = k->insert(KeyMapPair(mkey, it)).first; // maybe not ANSI compatible (?)
    }
    else return NULL;
  }
  return (*ki).second;
} /* FindItem */

HRESULT CAUSession::BecomeDB()
{ HRESULT hRet;
  if(!m_DB) CHKRET(g_CreateDB(&m_DB, false, L"DB/Session"));

  SAFEARRAY  sa, *psa=&sa;
  ASTR       sql(L"sess_CreateSession");
  VARIANT    var[2];
  UA1        tries=50;

  g_InitSA(&sa, var, 2);
  var[0].vt=VT_BSTR, var[0].bstrVal = SysAllocStringLen(NULL, 20);
  var[1].vt=VT_I4,   var[1].intVal  = m_Timeout;
  
  do
  { g_RandStr(var[0].bstrVal, 20); // TODO: replace random number generator [otherwise this may fail]
    hRet = m_DB->ExecuteONR(sql, L"@session_key,-session_id=3", &psa);
  } while(FAILED(hRet) && --tries);
  if(!tries) return E_FAIL;
  
  m_SessionID.Attach(var[0].bstrVal);
  IFS(m_DB->Output(sql=L"session_id", var)) m_DBID=var->intVal;
  else m_SessionID.Clear();
  return hRet;
} /* BecomeDB */

void CAUSession::DeleteSection(CAUSession::SectMap::iterator i)
{ KeyMap *km = (*i).second;
  KeyMap::iterator ki = km->begin();
  Item   *it;
  for(; ki != km->end(); ki++)
  { it = (*ki).second;
    VariantClear(&it->Var);
  }
  s_AllocCS.Lock();
  for(ki=km->begin(); ki!=km->end(); ki++) s_ItemAlloc.Delete((*ki).second);
  s_AllocCS.Unlock();
} /* DeleteSection */

void CAUSession::DeleteSection2(CAUSession::SectMap::iterator i)
{ if(m_DBID)
  { SAFEARRAY  sa, *psa=&sa;
    VARIANT    var[2];
    g_InitSA(&sa, var, 2);
    var[0].vt=VT_I4,   var[0].intVal=m_DBID;
    var[1].vt=VT_BSTR, var[1].bstrVal=(BSTR)(*i).first.InnerStr();
    m_DB->ExecuteONR(ASTR(L"sess_DeleteSection"), L"@session_id,@sect_name", &psa);
  }
  DeleteSection(i);
} /* DeleteSection2 */

HRESULT CAUSession::DeleteFromDB()
{ assert(m_DBID);
  SAFEARRAY sa, *psa=&sa;
  VARIANT   var;
  g_InitSA(&sa, &var);
  var.vt = VT_I4, var.intVal = m_DBID;
  return m_DB->ExecuteONR(ASTR(L"sess_DeleteSession"), L"@session_id", &psa);
} /* DeleteFromDB */

HRESULT CAUSession::ReadDBItem(CAUSession::Item *it, VARIANT *pvout)
{ assert(it && pvout && m_DBID && m_DB);
  ASTR       str(L"sess_GetValue");
  HRESULT    hRet;
  SAFEARRAY  sa, *psa;
  VARIANT    var[2];

  g_InitSA(&sa, var, 2);
  var[0].vt=var[1].vt=VT_I4, var[0].intVal=m_DBID, var[1].intVal=it->ID;
  CHKRET(m_DB->ExecuteONR(str, L"@session_id,+item_id,-binval=128/16,-textval=201", &psa));
  CHKRET(m_DB->Output(str=L"item_id", var));
  if(var->intVal==0) return E_ABORT;
  CHKRET(m_DB->Output(str=L"textval", pvout));
  if(pvout->vt==VT_NULL)
  { CHKRET(m_DB->Output(str=L"binval", var));
    if(var->vt&VT_ARRAY) memcpy(pvout, var->parray->pvData, sizeof(VARIANT));
    else hRet=E_ABORT;
    VariantClear(var);
  }
  return hRet;
} /* ReadDBItem */

HRESULT CAUSession::WriteDBItem(CAUSession::Item *it)
{ assert(it && it->ID && m_DBID && m_DB);
  HRESULT   hRet;
  SAFEARRAY sa, *psa, dsa;
  VARIANT   var[4];
  
  g_InitSA(&sa, var, 3);
  var[0].vt=var[1].vt=VT_I4, var[0].intVal=m_DBID, var[1].intVal=it->ID;
  
  if(it->Var.vt==VT_BSTR) var[2].vt=VT_NULL, var[3].vt=VT_BSTR, var[3].bstrVal=it->Var.bstrVal;
  else
  { g_InitSA(&dsa, &it->Var, 16);
    dsa.cbElements = 1;
    var[2].vt=VT_UI1|VT_ARRAY, var[3].vt=VT_NULL;
  }
  IFS(m_DB->ExecuteONR(ASTR(L"sess_PutValue"), L"@session_id,@item_id,@binval,@textval", &psa)) it->bDirty = false;
  return hRet;
} /* WriteDBItem */

HRESULT CAUSession::GetDBItem(ASTR &key, CAUSession::Item **ppit)
{ assert(ppit && m_DBID && m_DB);
  ASTR       str(L"sess_GetValue");
  HRESULT    hRet;
  SAFEARRAY  sa, *psa;
  VARIANT    var[3];

  g_InitSA(&sa, var, 2);
  var[0].vt=VT_I4, var[0].intVal=m_DBID, var[1].vt=VT_BSTR, var[1].bstrVal=(BSTR)key.InnerStr();
  CHKRET(m_DB->ExecuteONR(str, L"@session_id,@item_name,-item_id=3,-binval=128/16,-textval=201", &psa));
  CHKRET(m_DB->Output(str=L"item_id", var));
  if(var->intVal==0) return S_FALSE;

  CHKRET(m_DB->Output(str=L"textval", var+2));
  if(var[2].vt==VT_NULL)
  { CHKRET(m_DB->Output(str=L"binval", var+1));
    if(var[1].vt&VT_ARRAY) memcpy(var+2, var[1].parray->pvData, sizeof(VARIANT));
    else hRet=E_ABORT;
    VariantClear(var+1);
  }
  if(SUCCEEDED(hRet))
  { s_AllocCS.Lock();
    Item *it = s_ItemAlloc.New();
    s_AllocCS.Unlock();
    it->Var    = var[2];
    it->ID     = var[0].intVal;
    it->bDirty = false;
    (*(SplitKey(key))).second->insert(KeyMapPair(key, it));
  }
  return hRet;
} /* GetDBItem */

CAUSession::SectMap::iterator CAUSession::FindSection(const ASTR &key, bool autoInsert)
{ SectMap::iterator i = m_Sects.find(key);
  if(autoInsert && i==m_Sects.end())
    return m_Sects.insert(SectMapPair(ASTR(key), new KeyMap)).first; // this may not be ANSI compatible (?)
  return i;
} /* FindSection */

CAUSession::SectMap::iterator CAUSession::SplitKey(ASTR &key)
{ IA4 dot = key.IndexOf(L'.');
  if(dot == -1) return m_CurSect;
  SectMap::iterator i = FindSection(key.Substr(0, dot));
  key = key.Substr(dot+1);
  return i;
} /* SplitKey */

ACritSec CAUSession::s_AllocCS;
CBlockAlloc<CAUSession::Item> CAUSession::s_ItemAlloc(512, false, false);