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

/* ~(MODULES::VARARR, p'VarArray::Item
  <PRE>[propget] HRESULT Item([in] long nIndex, [out,retval] VARIANT *pvOut);</PRE>
  The Item property allows you to retrieve an item from the array by passing
  in the 0-based index.
)~ */
STDMETHODIMP CVarArray::get_Item(long nIndex, VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  if(nIndex<0 || (UA4)nIndex>=m_Max || (UA4)nIndex>=(UA4)m_Var.v.parray->rgsabound[0].cElements)
    return E_INVALIDARG;
  return VariantCopy(pvOut, (VARIANT*)m_Var.v.parray->pvData+nIndex);
} /* get_Item */

/* ~(MODULES::VARARR, p'VarArray::Size
  <PRE>[propget] HRESULT Size([out,retval] long *pnEls);</PRE>
  The Size property returns the number of items in the array.
)~ */
STDMETHODIMP CVarArray::get_Size(long *pnEls)
{ if(!pnEls) return E_POINTER;
  *pnEls = m_Max ? (long)m_Var.v.parray->rgsabound[0].cElements : 0;
  return S_OK;
} /* get_Size */

/* ~(MODULES::VARARR, p'VarArray::Capacity
  <PRE>[propget] HRESULT Capacity([out,retval] long *pnEls);
[propput] HRESULT Capacity([in] long nEls);</PRE>
  The Capacity property returns the number of slots in the array. The array will
  be grown as necessary, but you can assign a value to the Capacity property if
  you know the approximate size. It is not guaranteed that the capacity will
  actually be changed if you set the Capacity property, however.
)~ */
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
  
  if(m_Max != 0)
  { sa->rgsabound[0].cElements = m_Var.v.parray->rgsabound[0].cElements;
    memcpy(sa->pvData, m_Var.v.parray->pvData, sa->rgsabound[0].cElements*sizeof(VARIANT));
    m_Var.v.parray->rgsabound[0].cElements=0; // make sure data doesn't get wiped out
  }
  else sa->rgsabound[0].cElements = 0;

  VARIANT var;
  var.vt     = VT_ARRAY|VT_VARIANT;
  var.parray = sa;

  m_Var.Attach(var);
  m_Max = (UA4)nEls;
  return S_OK;
} /* put_Capacity */

/* ~(MODULES::VARARR, p'VarArray::Array
  <PRE>[propget] HRESULT Array([out,retval] VARIANT *pvOut);</PRE>
  The Array property returns the internal array as a VARIANT with a type of VT_ARRAY|VT_VARIANT.
  Reading from the Array property returns the internal array WITHOUT copying it, and then clears
  the internal array. This is more efficient than `ArrayCopy', but can only be read once, so don't
  throw away the value until you're done with it.
)~ */
STDMETHODIMP CVarArray::get_Array(VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  *pvOut = m_Var.Detach();
  m_Var.v.vt = VT_NULL;
  m_Max      = 0;
  return S_OK;
} /* get_Array */

/* ~(MODULES::VARARR, p'VarArray::ArrayCopy
  <PRE>[propget] HRESULT ArrayCopy([out,retval] VARIANT *pvOut);</PRE>
  The ArrayCopy property returns a copy of the the internal array each time it is read. Because it
  copies the array, it is much slower than `Array', but it exists in case you need it. The return
  value is a VARIANT of type VT_ARRAY|VT_VARIANT.
)~ */
STDMETHODIMP CVarArray::get_ArrayCopy(VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  return m_Var.CopyTo(pvOut);
} /* get_CopyArray */

/* ~(MODULES::VARARR, f'VarArray::Add
  <PRE>[propget] HRESULT Add([in] VARIANT *vEl);</PRE>
  The Add method adds an item to the array. If the new size of the array would be
  beyond the current capacity, the array is automatically grown.
)~ */
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
