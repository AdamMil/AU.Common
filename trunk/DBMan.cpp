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
#include "DB.h"

// CDBMan

CDBMan::CDBMan()
{ m_MaxShares = 8;
} /* CDBMan */

CDBMan::~CDBMan()
{ DBMap::iterator i = m_Map.begin();
  for(; i != m_Map.end(); i++)
  { DBList &list = *(*i).second;
    for(UA4 j=0; j<(UA4)list.size(); j++) list[j]->Release();
  }
} /* ~CDBMan */

STDMETHODIMP CDBMan::get_MaxSharing(long *pnShares)
{ if(!pnShares) return E_POINTER;
  AComLock lock(this);
  *pnShares = (long)m_MaxShares;
  return S_OK;
} /* get_MaxSharing */

STDMETHODIMP CDBMan::put_MaxSharing(long nShares)
{ if(nShares<0) return E_INVALIDARG;
  AComLock lock(this);
  m_MaxShares = (UA4)nShares;
  return S_OK;
} /* put_MaxSharing */

STDMETHODIMP CDBMan::CreateDB(BSTR sKey, BSTR sSect, IDB **ppDB)
{ assert(sKey && sSect);
  AComLock lock(this);
  AVAR     var;
  HRESULT  hRet;
  CHKRET(g_Config(sKey[0] ? sKey : ASTR(L"DB/Default"), ASTR(L"string"), sSect, var));
  return CreateDBRaw(var.ToBSTR(), ppDB);
} /* CreateDB */

STDMETHODIMP CDBMan::CreateDBRaw(BSTR sConnStr, IDB **ppDB)
{ if(!sConnStr || !ppDB) return E_POINTER;
  if(!sConnStr[0]) return E_INVALIDARG;
  
  AComPtr<IDB> db;
  HRESULT      hRet;
  if(m_MaxShares<2)
  { CHKRET(CREATE(DB, IDB, db));
    db->put_ConnString(sConnStr);
    db.CopyTo(ppDB);
    return S_OK;
  }

  AComLock lock(this);
  DBList      *list;
  UA4          i, len;
 
  DBMap::iterator it = m_Map.find(sConnStr);
  if(it == m_Map.end())
  { list = new DBList;
    m_Map.insert(std::pair<BSTR, DBList *>(SysAllocString(sConnStr), list));
  }
  else list = (*it).second;

  for(i=0,len=(UA4)list->size(); i<len; i++)
  { CDB *cdb = list->operator[]((int)i);
    if((UA4)cdb->m_dwRef-1 < m_MaxShares)
    { db = cdb;
      break;
    }
  }
  if(i == len)
  { CHKRET(CREATE(DB, IDB, db));
    list->push_back((CDB*)(IDB*)db);
    db->put_ConnString(sConnStr);
  }
  db.CopyTo(ppDB);
  return S_OK;
} /* CreateDBRaw */
