// SessionMan.cpp : Implementation of CSessionMan

#include "stdafx.h"
#include "SessionMan.h"


// CSessionMan

STDMETHODIMP CSessionMan::get_NumSessions(/*[out,retval]*/ long *pnSess)
{ return E_NOTIMPL;
} /* NumSessions */

STDMETHODIMP CSessionMan::get_Session(/*[in]*/ long nIndex, /*[out,retval]*/ IAUSession **ppSess)
{ return E_NOTIMPL;
} /* Session */

STDMETHODIMP CSessionMan::get_Current(/*[out,retval]*/ IAUSession **ppSess)
{ return E_NOTIMPL;
} /* Current */

STDMETHODIMP CSessionMan::Init(/*[in,defaultvalue("")]*/ BSTR sConfigKey, /*[in,defaultvalue(0)]*/ long nFlags)
{ return E_NOTIMPL;
} /* Init */

STDMETHODIMP CSessionMan::CreateSession(/*[in,defaultvalue(0)]*/ VARIANT_BOOL bPersist, /*[out,retval]*/ IAUSession **ppSess)
{ return E_NOTIMPL;
} /* CreateSession */

STDMETHODIMP CSessionMan::LoadSession(/*[in]*/ BSTR sID, /*[out,retval]*/ IAUSession **ppSess)
{ return E_NOTIMPL;
} /* LoadSession */

STDMETHODIMP CSessionMan::DropSession(/*[in]*/ BSTR sID)
{ return E_NOTIMPL;
} /* DropSession */

STDMETHODIMP CSessionMan::LockMan()
{ return E_NOTIMPL;
} /* LockMan */

STDMETHODIMP CSessionMan::UnlockMan()
{ return E_NOTIMPL;
} /* UnlockMan */
