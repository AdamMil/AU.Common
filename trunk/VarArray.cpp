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

// VarArray.cpp : Implementation of CVarArray

#include "stdafx.h"
#include "VarArray.h"


// CVarArray
STDMETHODIMP CVarArray::get_Item(long nIndex, VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  if(nIndex<0 || (UA4)nIndex>=m_Max || (UA4)nIndex>=(UA4)m_Var.v.parray->rgsabound[0].cElements)
    return E_INVALIDARG;
  return VariantCopy(pvOut, (VARIANT*)m_Var.v.parray->pvData+nIndex);
} /* get_Item */

STDMETHODIMP CVarArray::get_Size(long *pnEls)
{ if(!pnEls) return E_POINTER;
  *pnEls = m_Max ? (long)m_Var.v.parray->rgsabound[0].cElements : 0;
  return S_OK;
} /* get_Size */

STDMETHODIMP CVarArray::get_Capacity(long *pnEls)
{ if(!pnEls) return E_POINTER;
  *pnEls = (long)m_Max;
  return S_OK;
} /* get_Capacity */

STDMETHODIMP CVarArray::put_Capacity(long nEls)
{ if(nEls < (long)m_Max) return E_INVALIDARG;
  if(nEls == m_Max) return S_FALSE;
  SAFEARRAYBOUND bound = { nEls, 0 };
  SAFEARRAY     *sa    = SafeArrayCreate(VT_VARIANT, 1, &bound);
  if(sa==NULL) return E_FAIL;
  sa->rgsabound[0].cElements = m_Max ? m_Var.v.parray->rgsabound[0].cElements : 0;

  VARIANT var;
  var.vt     = VT_ARRAY|VT_VARIANT;
  var.parray = sa;

  m_Var.Attach(var);
  m_Max = (UA4)nEls;
  return S_OK;
} /* put_Capacity */

STDMETHODIMP CVarArray::get_Array(VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  return m_Var.CopyTo(pvOut);
} /* get_Array */

STDMETHODIMP CVarArray::Add(VARIANT vEl)
{ HRESULT hRet;
  UA4     len;
  if(m_Max==0)
  { len = 0;
    CHKRET(put_Capacity(8));
  }
  else
  { len = (UA4)m_Var.v.parray->rgsabound[0].cElements;
    if(len==m_Max) CHKRET(put_Capacity(m_Max*2));
  }
  CHKRET(VariantCopy((VARIANT*)m_Var.v.parray->pvData+len, &vEl));
  m_Var.v.parray->rgsabound[0].cElements++;
  return S_OK;
} /* Add */
