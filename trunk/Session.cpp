// Session.cpp : Implementation of CAUSession

#include "stdafx.h"
#include "Session.h"

// CAUSession

STDMETHODIMP CAUSession::get_Item(BSTR sKey, VARIANT *pvOut)
{ return E_NOTIMPL;
} /* get_Item */

STDMETHODIMP CAUSession::put_Item(BSTR sKey, VARIANT vData)
{ return E_NOTIMPL;
} /* put_Item */

STDMETHODIMP CAUSession::get_SessionID(BSTR *psID)
{ return E_NOTIMPL;
} /* get_SessionID */

STDMETHODIMP CAUSession::get_Section(BSTR *psSect)
{ return E_NOTIMPL;
} /* get_Section */

STDMETHODIMP CAUSession::put_Section(BSTR sSect)
{ return E_NOTIMPL;
} /* put_Section */

STDMETHODIMP CAUSession::get_Flags(long *pnFlags)
{ return E_NOTIMPL;
} /* get_Flags */

STDMETHODIMP CAUSession::put_Flags(long nFlags)
{ return E_NOTIMPL;
} /* put_Flags */

STDMETHODIMP CAUSession::get_Timeout(long *pnSecs)
{ return E_NOTIMPL;
} /* get_Timeout */

STDMETHODIMP CAUSession::put_Timeout(long nSecs)
{ return E_NOTIMPL;
} /* put_Timeout */

STDMETHODIMP CAUSession::get_IsManaged(VARIANT_BOOL *pbManaged)
{ return E_NOTIMPL;
} /* get_IsManaged */

STDMETHODIMP CAUSession::get_UsesDB(VARIANT_BOOL *pbDB)
{ return E_NOTIMPL;
} /* get_UsesDB */

STDMETHODIMP CAUSession::put_UsesDB(VARIANT_BOOL bDB)
{ return E_NOTIMPL;
} /* put_UsesDB */

STDMETHODIMP CAUSession::Clear(BSTR sSect)
{ return E_NOTIMPL;
} /* Clear */

STDMETHODIMP CAUSession::ClearAll()
{ return E_NOTIMPL;
} /* ClearAll */

STDMETHODIMP CAUSession::Persist(BSTR *psID)
{ return E_NOTIMPL;
} /* Persist */

STDMETHODIMP CAUSession::Revert()
{ return E_NOTIMPL;
} /* Revert */

STDMETHODIMP CAUSession::Save()
{ return E_NOTIMPL;
} /* Save */

STDMETHODIMP CAUSession::Delete()
{ return E_NOTIMPL;
} /* Delete */
