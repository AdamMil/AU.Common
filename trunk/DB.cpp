/* 
   This file is part of the AU.Common library, a set of ActiveX
   controls to aid in Web development.
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

// DB.cpp : Implementation of CDB

#include "stdafx.h"
#include "DB.h"
#include <mtx.h>

CDB::CDB()
{ m_nTimeout   = 30;
  m_CursorType = adOpenForwardOnly;
  m_LockType   = adLockReadOnly;
  m_bInit      = false;
}

STDMETHODIMP CDB::get_ConnSection(BSTR *pVal)
{ if(!pVal) return E_POINTER;
  *pVal = m_ConnSect.IsEmpty() ? SysAllocString(L"Default") : m_ConnSect.ToBSTR();
  return S_OK;
} /* get_ConnSection */

STDMETHODIMP CDB::put_ConnSection(BSTR newVal)
{ if(m_bInit && m_ConnSect != newVal) Close();
  m_ConnSect = newVal;
  return S_OK;
} /* put_ConnSection */

STDMETHODIMP CDB::get_ConnKey(BSTR *pVal)
{ if(!pVal) return E_POINTER;
  *pVal = m_ConnKey.IsEmpty() ? SysAllocString(L"Default") : m_ConnKey.ToBSTR();
  return S_OK;
} /* get_ConnKey */

STDMETHODIMP CDB::put_ConnKey(BSTR newVal)
{ if(m_bInit && m_ConnKey != newVal) Close();
  m_ConnKey = newVal;
  return S_OK;
} /* put_ConnKey */

STDMETHODIMP CDB::get_CursorType(int *pVal)
{ if(!pVal) return E_POINTER;
  *pVal = m_CursorType;
  return S_OK;
} /* get_CursorType */

STDMETHODIMP CDB::put_CursorType(int newVal)
{ m_CursorType = newVal;
  return S_OK;
} /* put_CursorType */

STDMETHODIMP CDB::get_LockType(int *pVal)
{ if(!pVal) return E_POINTER;
  *pVal = m_LockType;
  return S_OK;
} /* get_LockType */

STDMETHODIMP CDB::put_LockType(int newVal)
{ m_LockType = newVal;
  return S_OK;
} /* put_LockType */

STDMETHODIMP CDB::get_Timeout(long *pVal)
{ if(!pVal) return E_POINTER;
  *pVal = (long)m_LockType;
  return S_OK;
} /* get_Timeout */

STDMETHODIMP CDB::put_Timeout(long newVal)
{ if(newVal < 0) return E_INVALIDARG;
  m_nTimeout = (UA4)newVal;
  if(m_Cmd) return m_Cmd->put_CommandTimeout(newVal);
	return S_OK;
} /* put_Timeout */

STDMETHODIMP CDB::get_Connection(_Connection **pVal)
{ HRESULT hRet;
  if(!pVal)    return E_POINTER;
  if(!m_bInit) CHKRET(Init());
  m_Conn.CopyTo(pVal);
  return S_OK;
} /* get_Connection */

STDMETHODIMP CDB::get_Command(_Command **pVal)
{ HRESULT hRet;
  if(!pVal)    return E_POINTER;
  if(!m_bInit) CHKRET(Init());
  m_Cmd.CopyTo(pVal);
  return S_OK;
} /* get_Command */

STDMETHODIMP CDB::get_IsOpen(VARIANT_BOOL *pVal)
{ if(!pVal) return E_POINTER;
  *pVal = m_bInit ? VBTRUE : VBFALSE;
	return S_OK;
} /* get_IsOpen */

STDMETHODIMP CDB::Open()
{ return m_bInit ? S_OK : Init();
} /* Open */

STDMETHODIMP CDB::Close()
{ if(m_Cmd)  m_Cmd.Release();
  if(m_Conn) m_Conn->Close();
  m_bInit = false;
  return S_OK;
} /* Close */

STDMETHODIMP CDB::LockDB()
{ Lock();
  return S_OK;
} /* LockDB */

STDMETHODIMP CDB::UnlockDB()
{ Unlock();
  return S_OK;
} /* UnlockDB */

STDMETHODIMP CDB::NewCommand(_Command **pVal)
{ if(!pVal) return E_POINTER;
  AComPtr<ADOCommand> cmd;
  HRESULT hRet;
  VARIANT var;
  var.vt=VT_DISPATCH, var.pdispVal=NULL;

  if(!m_bInit) CHKRET(Init());
  CHKRET(CREATE(CADOCommand, IADOCommand, cmd));
  CHKRET(IQUERY(m_Conn, IDispatch, &var.pdispVal));
  IFS(cmd->put_ActiveConnection(var)) cmd.CopyTo(pVal);
  VariantClear(&var);
  return hRet;
} /* NewCommand */

STDMETHODIMP CDB::Execute(BSTR sSQL, SAFEARRAY(VARIANT) *aVals, ADORecordset **pRet)
{ if(!pRet) return ExecuteNR(sSQL, aVals);
  HRESULT hRet;
  CHKRET(StartExecute(sSQL));
  CHKRET(FillParams(aVals));
  return DoExecute(pRet);
} /* Execute */

STDMETHODIMP CDB::ExecuteNM(BSTR sSQL, BSTR sNames, SAFEARRAY(VARIANT) *aVals, ADORecordset **pRet)
{ if(!pRet) return ExecuteNMNR(sSQL, sNames, aVals);
  HRESULT hRet;
  CHKRET(StartExecute(sSQL));
  CHKRET(FillParamsNM(sNames, aVals));
  return DoExecute(pRet);
} /* ExecuteNM */

STDMETHODIMP CDB::ExecuteVal(BSTR sSQL, SAFEARRAY(VARIANT) *aVals, VARIANT *pRet)
{ if(!pRet) return E_POINTER;
  ARecordset rs;
  HRESULT    hRet;
  CHKRET(Execute(sSQL, aVals, &rs));
  return g_GetField(rs, 0, pRet);
} /* ExecuteVal */

STDMETHODIMP CDB::ExecuteNR(BSTR sSQL, SAFEARRAY(VARIANT) *aVals)
{ ARecordset rs;
  HRESULT    hRet;
  CHKRET(StartExecute(sSQL));
  CHKRET(FillParams(aVals));
  return m_Cmd->Execute(NULL, NULL, adExecuteNoRecords, &rs);
} /* ExecuteNR */

STDMETHODIMP CDB::ExecuteNMNR(BSTR sSQL, BSTR sParms, SAFEARRAY(VARIANT) *aVals)
{ ARecordset rs;
  HRESULT    hRet;
  CHKRET(StartExecute(sSQL));
  CHKRET(FillParamsNM(sParms, aVals));
  return m_Cmd->Execute(NULL, NULL, adExecuteNoRecords, &rs);
} /* ExecuteNMNR */

HRESULT CDB::Init()
{
} /* Init */

HRESULT CDB::StartExecute(BSTR sSQL)
{ HRESULT hRet;
  if(!m_bInit) CHKRET(Init());
  CHKRET(m_Cmd->put_CommandText(sSQL));
  CHKRET(m_Cmd->put_CommandType(wcspbrk(sSQL, L" \t\n") ? adCmdText : adCmdStoredProc));
  return hRet;
} /* StartExecute */

HRESULT CDB::DoExecute(ADORecordset **pRet)
{ ARecordset rs;
  HRESULT    hRet;
  VARIANT    var;

  CHKRET(CREATE(CADORecordset, IADORecordset, rs));

  var.vt=VT_DISPATCH, var.pdispVal=NULL;
  CHKRET(IQUERY(m_Cmd, IDispatch, &var.pdispVal));

  rs->put_CursorLocation(adUseClient);
  hRet = rs->Open(var, vtMissing, (CursorTypeEnum)m_CursorType, (LockTypeEnum)m_LockType, -1);
  if(SUCCEEDED(hRet) && m_LockType==adLockReadOnly && (m_CursorType==adOpenForwardOnly || m_CursorType==adOpenStatic))
    rs->putref_ActiveConnection(NULL);
  VariantClear(&var);
  ResetDefaults();
  if(SUCCEEDED(hRet)) rs.CopyTo(pRet);
  return hRet;
} /* DoExecute */

void CDB::ResetDefaults()
{ m_CursorType = adOpenForwardOnly;
  m_LockType   = adLockReadOnly;
  put_Timeout(30);
} /* ResetDefaults */
