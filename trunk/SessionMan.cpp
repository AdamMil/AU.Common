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
