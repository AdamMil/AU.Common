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

STDMETHODIMP CDB::get_Connection(ADOConnection **pVal)
{ HRESULT hRet;
  if(!pVal)    return E_POINTER;
  if(!m_bInit) CHKRET(Init());
  m_Conn.CopyTo(pVal);
  return S_OK;
} /* get_Connection */

STDMETHODIMP CDB::get_Command(ADOCommand **pVal)
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

STDMETHODIMP CDB::get_Output(BSTR sParam, VARIANT *pvOut)
{ if(!m_Cmd) return E_UNEXPECTED;
  if(!pvOut) return E_POINTER;
  AComPtr<ADOParameters> parms;
  AComPtr<ADOParameter>  parm;
  HRESULT hRet;
  VARIANT var;

  var.vt=VT_BSTR, var.bstrVal=sParam;
  CHKRET(m_Cmd->get_Parameters(&parms));
  CHKRET(parms->get_Item(var, &parm))
  return parm->get_Value(pvOut);
} /* get_Output */

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

STDMETHODIMP CDB::NewCommand(ADOCommand **pVal)
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

STDMETHODIMP CDB::ExecuteO(BSTR sSQL, BSTR sNames, SAFEARRAY(VARIANT) *aVals, ADORecordset **pRet)
{ if(!pRet) return ExecuteONR(sSQL, sNames, aVals);
  HRESULT hRet;
  CHKRET(StartExecute(sSQL));
  CHKRET(FillParamsO(sNames, aVals));
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

STDMETHODIMP CDB::ExecuteONR(BSTR sSQL, BSTR sParms, SAFEARRAY(VARIANT) *aVals)
{ ARecordset rs;
  HRESULT    hRet;
  CHKRET(StartExecute(sSQL));
  CHKRET(FillParamsO(sParms, aVals));
  return m_Cmd->Execute(NULL, NULL, adExecuteNoRecords, &rs);
} /* ExecuteNMNR */

HRESULT CDB::Init()
{ HRESULT hRet=S_OK;
  if(m_Conn==NULL) CHKRET(CREATE(CADOConnection, IADOConnection, m_Conn));
  if(m_Cmd ==NULL)
  { CHKRET(CREATE(CADOCommand, IADOCommand, m_Cmd));
    m_Cmd->put_CommandTimeout(m_nTimeout);
  }
  
  long state;
  CHKRET(m_Conn->get_State(&state));
  if(state == adStateClosed)
  { AVAR var;
    CHKRET(g_Config(m_ConnKey.IsEmpty() ? ASTR(L"Default") : m_ConnKey, ASTR(L"string"), m_ConnSect, var));
    CHKRET(m_Conn->Open(var.ToBSTR(), NULL, NULL, -1));

    var.Clear();
    CHKRET(IQUERY(m_Conn, IDispatch, &var.v.pdispVal))
    var.v.vt = VT_DISPATCH;
    hRet = m_Cmd->put_ActiveConnection(var);
  }
  m_bInit = SUCCEEDED(hRet);
  return hRet;
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
  hRet = rs->Open(var, g_vMissing, (CursorTypeEnum)m_CursorType, (LockTypeEnum)m_LockType, -1);
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

HRESULT CDB::FillParams(SAFEARRAY **aVals)
{ AComPtr<ADOParameters> parms;
  AComPtr<ADOParameter>  parm;
  AVAR       var;
  SAFEARRAY *array = aVals && *aVals ? *aVals : NULL;
  HRESULT    hRet;
  
  CHKRET(m_Cmd->get_Parameters(&parms));
  if(array == NULL || FAILED(hRet=parms->Refresh()))
  { long len=0;
    parms->get_Count(&len);
    var.SetI4(0);
    while(len--) parms->Delete(var);
    if(array == NULL) return S_FALSE;
  }

  VARIANT *vars = (VARIANT*)array->pvData;
  long vi, nvars = array->rgsabound[0].cElements, nparms=0;

  if(SUCCEEDED(hRet)) // Refresh() succeeded
  { ParameterDirectionEnum e;
    parms->get_Count(&nparms);
    var.SetI4(0);
    for(vi=0; vi<nvars && var.v.intVal<nparms; var.v.intVal++)
    { CHKRET(parms->get_Item(var, &parm));
      parm->get_Direction(&e);
      if(e&1) parm->put_Value(vars[vi++]); // e&1 is nonzero if it's an input parameter
      parm.Release();
    }
  }
  else
  { static DataTypeEnum dbTypes[VT_UINT] =
    { adEmpty,   adVarChar,   adSmallInt, adInteger, adSingle,  adDouble,   adCurrency, adDate,
      adVarChar, adIDispatch, adError,    adBoolean, adVariant, adIUnknown, adDecimal,  adTinyInt,
      adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adBigInt, adUnsignedBigInt, adInteger,
      adUnsignedInt
    };
    VARIANT *pvar;
    long     size, type;
    VARTYPE  vt;
    
    for(vi=0; vi<nvars; vi++)
    { pvar  = (vars[vi].vt&VT_BYREF) ? vars[vi].pvarVal : vars+vi, vt = pvar->vt;
      if(vt > VT_UINT) return E_INVALIDARG;
      type = dbTypes[vt];
      if(vt & VT_ARRAY) type &= 0x2000;
      size = (vt==VT_BSTR ? SysStringLen(pvar->bstrVal) : 1);

      CHKRET(m_Cmd->CreateParameter((BSTR)ASTR::EmptyBSTR, (DataTypeEnum)type, adParamInput, size, *pvar, &parm));
      CHKRET(parms->Append(parm));
      parm.Release();
    }      
  }
  return hRet;
} /* FillParams */


HRESULT CDB::FillParamsO(BSTR sParms, SAFEARRAY **aVals)
{ 
} /* FillParamsO */
