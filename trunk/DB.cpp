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

// DB.cpp : Implementation of CDB

#include "stdafx.h"
#include "DB.h"

CDB::CDB()
{ m_nTimeout   = 30;
  m_CursorType = adOpenForwardOnly;
  m_LockType   = adLockReadOnly;
  m_ConnStr    = NULL;
  m_bInit      = false;
} /* CDB */

CDB::~CDB()
{ SysFreeString(m_ConnStr);
} /* ~CDB */

/* ~(MODULES::DB, p'DB::ConnSection
  <PRE>
  [propget] HRESULT ConnSection([out,retval] BSTR *psSect);
  [propput] HRESULT ConnSection([in] BSTR sSect);</PRE>
  The ConnSection property sets the `Config' Section used to retrieve the ADO connection string.
  The combination of the ConnSection and ConnKey properties determines what ADO connection string
  will be used, unless `ConnString' is set.
  If the DB object is open when this property is changed, it will first be closed.
  The ConnSection defaults to "Default".
)~ */
STDMETHODIMP CDB::get_ConnSection(BSTR *psSect)
{ if(!psSect) return E_POINTER;
  AComLock lock(this);
  *psSect = m_ConnSect.IsEmpty() ? SysAllocString(L"Default") : m_ConnSect.ToBSTR();
  return S_OK;
} /* get_ConnSection */

STDMETHODIMP CDB::put_ConnSection(BSTR sSect)
{ AComLock lock(this);
  if(m_bInit && m_ConnSect != sSect) Close();
  m_ConnSect = sSect;
  return S_OK;
} /* put_ConnSection */

/* ~(MODULES::DB, p'DB::ConnKey
  <PRE>
  [propget] HRESULT ConnKey([out,retval] BSTR *psKey);
  [propput] HRESULT ConnKey([in] BSTR sKey);</PRE>
  The ConnSection property sets the `Config' Key used to retrieve the ADO Connection string.
  The combination of the ConnSection and ConnKey properties determines what ADO connection string
  will be used, unless `ConnString' is set.
  If the DB object is open when this property is changed, it will first be closed.
  The ConnKey defaults to "DB/Default".
)~ */
STDMETHODIMP CDB::get_ConnKey(BSTR *psKey)
{ if(!psKey) return E_POINTER;
  AComLock lock(this);
  *psKey = m_ConnKey.IsEmpty() ? SysAllocString(L"DB/Default") : m_ConnKey.ToBSTR();
  return S_OK;
} /* get_ConnKey */

STDMETHODIMP CDB::put_ConnKey(BSTR sKey)
{ AComLock lock(this);
  if(m_bInit && m_ConnKey != sKey) Close();
  m_ConnKey = sKey;
  return S_OK;
} /* put_ConnKey */

/* ~(MODULES::DB, p'DB::ConnString
  <PRE>
  [propget] HRESULT ConnString([out,retval] BSTR *psStr);
  [propput] HRESULT ConnString([in] BSTR sStr);</PRE>
  The ConnString property overrides the `ConnSection' and `ConnKey' properties. It
  allows an arbitrary connection string to be set, rather than a predefined one
  from the `Config' object. Setting this to an empty string will disable it and
  allow the Section/Key properties to take effect again. If the DB object is open
  when this property is changed, it will first be closed.
)~ */
STDMETHODIMP CDB::get_ConnString(BSTR *psStr)
{ if(!psStr) return E_POINTER;
  AComLock lock(this);
  if(m_ConnStr==NULL)
  { AVAR    var;
    HRESULT hRet;
    CHKRET(g_Config(m_ConnKey.IsEmpty() ? ASTR(L"DB/Default") : m_ConnKey, ASTR(L"string"), m_ConnSect, var));
    *psStr = var.Detach().bstrVal;
  }
  else *psStr = SysAllocString(m_ConnStr);
  return S_OK;
} /* get_ConnString */

STDMETHODIMP CDB::put_ConnString(BSTR sStr)
{ AComLock lock(this);
  if(m_bInit) Close();
  SysFreeString(m_ConnStr);
  m_ConnStr = sStr&&sStr[0] ? SysAllocString(sStr) : NULL;
  return S_OK;
} /* put_ConnString */

/* ~(MODULES::DB, p'DB::CursorType
  <PRE>
  [propget] HRESULT CursorType([out,retval] int *pnType);
  [propput] HRESULT CursorType([in] int nType);</PRE>
  This property determines what the ADO cursor type will be for the next Execute* call. It
  accepts and returns values from the ADO cursor type enum. This value will be reset to
  the default after the next Execute* call (successful or not). This defaults to
  adOpenForwardOnly.
)~ */
STDMETHODIMP CDB::get_CursorType(int *pnType)
{ if(!pnType) return E_POINTER;
  AComLock lock(this);
  *pnType = m_CursorType;
  return S_OK;
} /* get_CursorType */

STDMETHODIMP CDB::put_CursorType(int nType)
{ AComLock lock(this);
  m_CursorType = nType;
  return S_OK;
} /* put_CursorType */

/* ~(MODULES::DB, p'DB::LockType
  <PRE>
  [propget] HRESULT LockType([out,retval] int *pnType);
  [propput] HRESULT LockType([in] int nType);</PRE>
  This property determines what the ADO lock type will be for the next Execute* call. It
  accepts and returns values from the ADO lock type enum. This value will be reset to
  the default after the next Execute* call (successful or not). This defaults to
  adLockReadOnly.
)~ */
STDMETHODIMP CDB::get_LockType(int *pnType)
{ if(!pnType) return E_POINTER;
  AComLock lock(this);
  *pnType = m_LockType;
  return S_OK;
} /* get_LockType */

STDMETHODIMP CDB::put_LockType(int nType)
{ AComLock lock(this);
  m_LockType = nType;
  return S_OK;
} /* put_LockType */

/* ~(MODULES::DB, p'DB::Timeout
  <PRE>
  [propget] HRESULT Timeout([out,retval] long *pnTimeout);
  [propput] HRESULT Timeout([in] long nTimeout);</PRE>
  This property determines what the timeout in seconds will be for the next Execute* command.
  You can set this to 0 to disable the timeout entirely. This value will be reset to the
  default after the next Execute* call (successful or not). This defaults to 30 seconds.
)~ */
STDMETHODIMP CDB::get_Timeout(long *pnTimeout)
{ if(!pnTimeout) return E_POINTER;
  AComLock lock(this);
  *pnTimeout = (long)m_LockType;
  return S_OK;
} /* get_Timeout */

STDMETHODIMP CDB::put_Timeout(long nTimeout)
{ if(nTimeout < 0) return E_INVALIDARG;
  AComLock lock(this);
  m_nTimeout = (UA4)nTimeout;
  if(m_Cmd) return m_Cmd->put_CommandTimeout(nTimeout);
  return S_OK;
} /* put_Timeout */

/* ~(MODULES::DB, p'DB::Connection
  <PRE>[propget] HRESULT Connection([out,retval] ADOConnection *ppConn);</PRE>
  This property returns a reference to the internal ADOConnection object. It is
  recommended that not you use this property unless you really need to.
)~ */
STDMETHODIMP CDB::get_Connection(ADOConnection **ppConn)
{ HRESULT hRet;
  if(!ppConn)  return E_POINTER;
  AComLock lock(this);
  if(!m_bInit) CHKRET(Init());
  m_Conn.CopyTo(ppConn);
  return S_OK;
} /* get_Connection */

/* ~(MODULES::DB, p'DB::Command
  <PRE>[propget] HRESULT Command([out,retval] ADOCommand *ppCmd);</PRE>
  This property returns a reference to the internal ADOCommand object. It is
  recommended that not you use this property unless you really need to.
)~ */
STDMETHODIMP CDB::get_Command(ADOCommand **ppCmd)
{ HRESULT hRet;
  if(!ppCmd)   return E_POINTER;
  AComLock lock(this);
  if(!m_bInit) CHKRET(Init());
  m_Cmd.CopyTo(ppCmd);
  return S_OK;
} /* get_Command */

/* ~(MODULES::DB, p'DB::IsOpen
  <PRE>[propget] HRESULT IsOpen([out,retval] VARIANT_BOOL *pbOpen);</PRE>
  This property returns true (nonzero) if there is currently an active database
  connection and false (zero) otherwise. If the connection was terminated
  abnormally, this condition will not be caught. The connection will be open,
  but calls will fail. If that happens, you should call Close() to reset the
  internal status.
)~ */
STDMETHODIMP CDB::get_IsOpen(VARIANT_BOOL *pbOpen)
{ if(!pbOpen) return E_POINTER;
  AComLock lock(this);
  *pbOpen = m_bInit ? VBTRUE : VBFALSE;
  return S_OK;
} /* get_IsOpen */

/* ~(MODULES::DB, f'DB::Open
  <PRE>HRESULT Open();</PRE>
  The Open method opens the connection to the database if it is not open already.
  Generally, you will not need to call this method as the DB object will open
  itself when necessary.
)~ */
STDMETHODIMP CDB::Open()
{ AComLock lock(this);
  return m_bInit ? S_OK : Init();
} /* Open */

/* ~(MODULES::DB, f'DB::Close
  <PRE>HRESULT Close();</PRE>
  The Close method closes the connection to the database if it is open. The DB object
  does not automatically close the connection when it is destroyed. However, you
  shouldn't need to call this function, as the ADOConnection object will close the
  connection when its last reference is released.
)~ */
STDMETHODIMP CDB::Close()
{ AComLock lock(this);
  if(m_Cmd)  m_Cmd.Release();
  if(m_Conn) m_Conn->Close();
  m_bInit = false;
  return S_OK;
} /* Close */

/* ~(MODULES::DB, f'DB::LockDB
  <PRE>HRESULT LockDB();</PRE>
  The LockDB method is used in conjunction with the `UnlockDB' method to lock the
  state of the database if you will be performing a multi-call operation and the
  database object is shared between threads. An example of a "multi-call operation"
  is setting some of the properties (like CursorType, LockType, Timeout, etc) and
  then calling Execute. To prevent another thread from calling Execute between your
  calls to set the properties (causing them to be reset after the Execute finishes),
  you should call this method before the batch of commands and call `UnlockDB'
  afterwards. You do not need to use LockDB/UnlockDB if you are performing calls that
  are not related, such as two different Executes. Failure to call `UnlockDB' when
  you are done may cause other threads to lock up. If multiple threads exhibit this
  bug, it could cause an application deadlock.
)~ */

STDMETHODIMP CDB::LockDB()
{ Lock();
  return S_OK;
} /* LockDB */

/* ~(MODULES::DB, f'DB::UnlockDB
  <PRE>HRESULT UnlockDB();</PRE>
  The UnlockDB method is used in conjunction with the `LockDB' method. See the `LockDB'
  method for more information.
)~ */
STDMETHODIMP CDB::UnlockDB()
{ Unlock();
  return S_OK;
} /* UnlockDB */

/* ~(MODULES::DB, f'DB::NewCommand
  <PRE>HRESULT NewCommand([out,retval] ADOCommand *ppCmd);</PRE>
  The NewCommand method creates a new ADOCommand object and attaches it to the
  internal ADOConnection object, using its ActiveConnection property. It is
  recommended that you don't call this function unless you really need to. You might
  want to just create another DB object instead. If multiple threads are going to be
  using the DB object, you probably want to use `LockDB' to wrap the entire lifetime
  of the new ADOCommand object in a locked section (from before the NewCommand call
  until the created ADOCommand object is destroyed). Calling stored procedures with
  output parameters may fail if multiple ADOCommand objects are using the same
  ADOConnection object.
)~ */
STDMETHODIMP CDB::NewCommand(ADOCommand **ppCmd)
{ if(!ppCmd) return E_POINTER;
  AComPtr<ADOCommand> cmd;
  HRESULT hRet;
  VARIANT var;
  var.vt=VT_DISPATCH, var.pdispVal=NULL;

  AComLock lock(this);
  if(!m_bInit) CHKRET(Init());
  CHKRET(CREATE(CADOCommand, IADOCommand, cmd));
  CHKRET(IQUERY(m_Conn, IDispatch, &var.pdispVal));
  IFS(cmd->put_ActiveConnection(var)) cmd.CopyTo(ppCmd);
  VariantClear(&var);
  return hRet;
} /* NewCommand */

/* ~(MODULES::DB, f'DB::Execute
  <PRE>HRESULT Execute([in] BSTR sSQL, [in] SAFEARRAY(VARIANT) *aVals, [out,retval] ADORecordset **ppRS);</PRE>
  The Execute command is a simple method to execute an SQL query. The SQL is passed
  in 'sSQL'. If 'sSQL' contains any whitespace, it is assumed to be a
  query. Otherwise, it is assumed to be the name of a stored procedure. A recordset
  is then returned containing the results. If the results will not be used, using the
  Execute*NR functions will give optimal performance, as the relatively large
  overhead of constructing a recordset can be avoided. This method supports '?'
  replacement in the query. The replacement parameters are taken from the 'aVals'
  array. If no parameters are necessary, an empty array or just NULL can be passed
  for the value of 'aVals'. An example of '?' replacement is the following:
  <PRE>oRS = oDB.Execute("SELECT * FROM table WHERE col1=? AND col2 LIKE ?", 5, sFilter);</PRE>
  Each '?' is replaced by the respective value in the parameter array. The '?'
  replacement is done in the ODBC driver and is safer and faster than trying to
  interpolate the values into the SQL string. If 'sSQL' is the name of a stored
  procedure, then the parameters should be the input parameters to the stored
  procedure, in the correct order. Output parameters are not supported by this
  function, but they are by the ExecuteO* functions. '?' replacement also helps prevent
  cases where malicious SQL can be inserted into the string.  For instance:
  <PRE>oDB.Execute("SELECT * FROM table WHERE col1="+sInput);</PRE>
  If 'sInput' is "5; DROP DATABASE accounting", you could have trouble. Instead of
  attempting to clean all the input, using '?' replacement will eliminate this
  danger. This means, though, that the following will not work:
  <PRE>oDB.Execute("SELECT ? FROM table WHERE col1=?", "col5,col9", 5);</PRE>
  The replacements cannot be arbitrary SQL, only values. There are also performance
  benefits to using '?' replacement. The SQL engine will be able to cache the query
  better, and may not have to recompile the query again on the next call. This is
  because the following two queries can be considered to be indentical by the query
  compiler, and the compiled version can be cached:
  <PRE>
  oDB.Execute("SELECT * FROM table WHERE col1=?", 5);
  oDB.Execute("SELECT * FROM table WHERE col1=?", 17);</PRE>
  However, inserting a literal 5 or 17 into query will likely render the query cache
  useless unless the exact same values are used again. At the same time, there is a
  performance penalty associated with using '?' replacement. The DB object will have
  to make 2 round trips to the engine instead of one. The first round trip asks the
  engine to parse the query and return a list of parameters.  The data is used to
  fill in the parameters, and then the query is executed. To avoid this penalty and
  get the highest performance, consider using the ExecuteO* functions.
)~ */
STDMETHODIMP CDB::Execute(BSTR sSQL, SAFEARRAY(VARIANT) *aVals, ADORecordset **ppRS)
{ if(!ppRS) return E_POINTER;
  HRESULT  hRet;
  AComLock lock(this);
  CHKRT2(StartExecute(sSQL), Done);
  CHKRT2(FillParams(aVals), Done);
  hRet = DoExecute(ppRS);
  Done:
  ResetDefaults();
  return hRet;
} /* Execute */

/* ~(MODULES::DB, f'DB::ExecuteO
  <PRE>HRESULT ExecuteO([in] BSTR sSQL, [in] BSTR sParms,
                 [in] SAFEARRAY(VARIANT) *aVals, [out,retval] ADORecordset **ppRS);</PRE>
  The ExecuteO command is a more flexible method than `Execute' to execute a SQL
  query.  This method is similar to `Execute', but it does not support '?'
  replacement. Instead it supports a more powerful method that allows the use of
  input/output/inout parameters. This method has all of the benefits associated with
  '?' replacement described in the help for `Execute'.  Output and in/out parameters
  can be accessed after the query has completed by using the `Output' method. Remember
  that This also avoids the extra round trip to the database engine. With this
  function, a list of parameter names is passed in 'sParms'.  'sParms' is in the
  following format:
  <PRE>
  Pname[=[type][/size]],Pname[=[type][/size]],Pname[=[type][/size]]...

  P is assumed to be a prefix, which is one of the following:
    @ indicates an input parameter
    - indicates an output parameter
    + indicates an in/out parameter

  type is assumed to be a type that matches one of the adDataTypeEnum values.
  the type should be specified for all output parameters. The type can also
  be specified for input or in/out parameters, but this is generally not
  recommended. For input and in/out parameters, the DB object will attempt to
  guess the type from the value passed in. However, you may want the type
  converted into something else by ADO.
  
  size can be used to override the default size, if you want to do that. This
  is problematic for in/out parameters that are variable-width, because the size
  would specify both the input and output sizes, which probably won't be the
  same. It's best to stay away from in/out parameters except for simple scalar
  types.
  
  An example 'sParms' string is the following, which has two input parameters
  and an output parameter, which is of type adBoolean:
  @login,@pass,-allowed=11</PRE>
  
  The parameters passed to ExecuteO are expected to be in the order listed in
  'sParms'. Output parameters aren't passed, of course. An example usage may be
  the following:
  <PRE>
  oDB.ExecuteONR("users_CheckLogin", "@login,@pass,-allowed=11", sLogin, sPass);
  if(oDB.Output("allowed")) { let user in }
  </PRE>
  
  When using queries, instead of stored procedures, here is an example:
  <PRE>oRS = oDB.ExecuteO("SELECT * FROM t1 WHERE col1=@1 and col2=@2 and col3=@1",
                   "@1,@2", 5, "hello");</PRE>
  Queries can't have output or in/out parameters, for obvious reasons. However, there
  is still a performance benefit from using ExecuteO* methods instead of the standard
  Execute* ones. Note how parameters can be reused with the ExecuteO* functions. The
  @1 parameter is passed once and used twice. Also, in the query text, the parameters
  are always prefixed with @.
)~ */
STDMETHODIMP CDB::ExecuteO(BSTR sSQL, BSTR sParms, SAFEARRAY(VARIANT) *aVals, ADORecordset **ppRS)
{ if(!ppRS) return E_POINTER;
  HRESULT  hRet;
  AComLock lock(this);
  CHKRT2(StartExecute(sSQL), Done);
  CHKRT2(FillParamsO(sParms, aVals), Done);
  hRet = DoExecute(ppRS);
  Done:
  ResetDefaults();
  return hRet;
} /* ExecuteO */

/* ~(MODULES::DB, f'DB::ExecuteVal
  <PRE>HRESULT ExecuteVal([in] BSTR sSQL, [out,retval] VARIANT *pvOut);</PRE>
  The ExecuteVal method is a shortcut. It works by calling `Execute' and
  returning the value of the first column in the first row of whatever the
  recordset is.
)~ */
STDMETHODIMP CDB::ExecuteVal(BSTR sSQL, SAFEARRAY(VARIANT) *aVals, VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  ARecordset rs;
  HRESULT    hRet;
  AComLock lock(this);
  CHKRET(Execute(sSQL, aVals, &rs));
  return g_GetField(rs, 0, pvOut);
} /* ExecuteVal */

/* ~(MODULES::DB, f'DB::ExecuteNR
  <PRE>HRESULT ExecuteNR([in] BSTR sSQL, [in] SAFEARRAY(VARIANT) *aVals);</PRE>
  The ExecuteNR method acts like `Execute', except that it doesn't create
  a recordset, avoiding the relatively large overhead associated with that
  process. This can be used with statements that genuinely return no results,
  such as UPDATE or INSERT, and also with statements that usually do return
  results, like SELECT. In the latter case, the results will be ignored.
)~ */
STDMETHODIMP CDB::ExecuteNR(BSTR sSQL, SAFEARRAY(VARIANT) *aVals)
{ ARecordset rs;
  HRESULT    hRet;
  AComLock   lock(this);
  CHKRT2(StartExecute(sSQL), Done);
  CHKRT2(FillParams(aVals), Done);
  hRet = m_Cmd->Execute(NULL, NULL, adExecuteNoRecords, &rs);
  Done:
  ResetDefaults();
  return hRet;
} /* ExecuteNR */

/* ~(MODULES::DB, f'DB::ExecuteONR
  <PRE>HRESULT ExecuteONR([in] BSTR sSQL, [in] BSTR sParms, [in] SAFEARRAY(VARIANT) *aVals);</PRE>
  The ExecuteONR method acts like `ExecuteO', except that it doesn't create
  a recordset, avoiding the relatively large overhead associated with that
  process. This can be used with statements that genuinely return no results,
  such as UPDATE or INSERT, and also with statements that usually do return
  results, like SELECT. In the latter case, the results will be ignored.
)~ */
STDMETHODIMP CDB::ExecuteONR(BSTR sSQL, BSTR sParms, SAFEARRAY(VARIANT) *aVals)
{ ARecordset rs;
  HRESULT    hRet;
  AComLock   lock(this);
  CHKRT2(StartExecute(sSQL), Done);
  CHKRT2(FillParamsO(sParms, aVals), Done);
  hRet = m_Cmd->Execute(NULL, NULL, adExecuteNoRecords, &rs);
  Done:
  ResetDefaults();
  return hRet;
} /* ExecuteNMNR */

/* ~(MODULES::DB, f'DB::Output
  <PRE>HRESULT Output([in] BSTR sParam, [out,retval] VARIANT *pvOut);</PRE>
  This method is used to retrieve output parameters from one of the ExecuteO*
  methods. It takes the name of the parameter (without any leading prefix, such as "@")
  and returns the value. If the DB object is being used concurrently by multiple threads,
  `LockDB' should be used to lock the object from before the ExecuteO* call until all
  the output parameters have been retrieved, as they will be overwritten if another
  thread calls Execute*.
)~ */
STDMETHODIMP CDB::Output(BSTR sParam, VARIANT *pvOut)
{ if(!m_Cmd) return E_UNEXPECTED;
  if(!pvOut) return E_POINTER;
  AComPtr<ADOParameters> parms;
  AComPtr<ADOParameter>  parm;
  HRESULT  hRet;
  VARIANT  var;
  AComLock lock(this);
  
  var.vt=VT_BSTR, var.bstrVal=sParam;
  CHKRET(m_Cmd->get_Parameters(&parms));
  CHKRET(parms->get_Item(var, &parm));
  return parm->get_Value(pvOut);
} /* get_Output */

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
  { BSTR connstr;
    AVAR var;
    
    if(m_ConnStr==NULL)
    { CHKRET(g_Config(m_ConnKey.IsEmpty() ? ASTR(L"DB/Default") : m_ConnKey, ASTR(L"string"), m_ConnSect, var));
      connstr = var.Detach().bstrVal;
    }
    else connstr=m_ConnStr;
    hRet = m_Conn->Open(connstr, NULL, NULL, -1);
    if(m_ConnStr==NULL) SysFreeString(connstr);
    CHKRET(hRet);

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
      if(e&1) parm->put_Value(vars[vi++]);
      parm.Release();
    }
  }
  else
  { VARIANT *pvar;
    long     size, type;
    VARTYPE  vt;
    
    for(vi=0; vi<nvars; vi++)
    { pvar  = (vars[vi].vt&VT_BYREF) ? vars[vi].pvarVal : vars+vi, vt = pvar->vt;
      if(vt > VT_UINT) return E_INVALIDARG;
      type = m_dbTypes[vt];
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
{ AComPtr<ADOParameters> parms;
  AComPtr<ADOParameter>  parm;
  AVAR       var;
  SAFEARRAY *array = aVals && *aVals ? *aVals : NULL;
  ASTRList   list = ASTR(sParms).Split(L",");
  ASTR       name;
  VARIANT   *pvar, *vars;
  WCHAR     *pc, *p2;
  long       len=0, size, type;
  UA4        li, vi, nvars=0;
  HRESULT    hRet;
  VARTYPE    vt;
  
  for(li=0; li<list.Length(); li++)
  { if(list[li].Length()<2) return E_INVALIDARG;
    if(list[li][0]==L'@' || list[li][0]==L'+') nvars++;
  }
  if(nvars && (!array || array->rgsabound[0].cElements != nvars)) return E_INVALIDARG;

  CHKRET(m_Cmd->get_Parameters(&parms));
  parms->get_Count(&len);
  var.SetI4(0);
  while(len--) parms->Delete(var);
  if(array == NULL) return S_FALSE;
  vars = (VARIANT*)array->pvData;
  
  for(li=vi=0; li<list.Length(); li++)
  { ASTR &s = list[li], name = (WCHAR*)s+1, fc=s[0];
    if(fc==L'@' || fc==L'+')
    { pvar = (vars[vi].vt&VT_BYREF) ? vars[vi].pvarVal : vars+vi, vt = pvar->vt;
      size = (vt==VT_BSTR ? SysStringLen(pvar->bstrVal) : 1);
      if(vt > VT_UINT) return E_INVALIDARG;
      type = m_dbTypes[vt];
      if(vt & VT_ARRAY) type &= 0x2000;
      vi++;
    }
    else pvar=NULL, type=0, size=1;
    if((pc=wcschr(name, L'=')))
    { p2=wcschr(++pc, L'/');
      if(p2 != pc)   type=_wtol(pc);
      if(p2 != NULL) size=_wtol(p2+1);
    }
    else if(fc==L'-') return E_INVALIDARG;
    CHKRET(m_Cmd->CreateParameter(name, (DataTypeEnum)type,
                                  fc==L'@' ? adParamInput : fc==L'+' ? adParamInputOutput : adParamOutput,
                                  size, pvar ? *pvar : g_vMissing, &parm));
    CHKRET(parms->Append(parm));
    parm.Release();
  }
  return hRet;
} /* FillParamsO */

const DataTypeEnum CDB::m_dbTypes[VT_UINT] =
{ adEmpty,   adVarChar,   adSmallInt, adInteger, adSingle,  adDouble,   adCurrency, adDate,
  adVarChar, adIDispatch, adError,    adBoolean, adVariant, adIUnknown, adDecimal,  adTinyInt,
  adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adBigInt, adUnsignedBigInt, adInteger,
  adUnsignedInt
};
