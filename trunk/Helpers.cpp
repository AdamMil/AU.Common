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

#include "stdafx.h"
#include "Helpers.h"

/*** ASTR implementation ***/
ASTR & ASTR::Format(const WCHAR *ctl, ...)
{ std::va_list list;
  UA4   len;
  WCHAR buf[4096];
  va_start(list, ctl);
  len = (UA4)vswprintf(buf, ctl, list);
  va_end(list);
  Set(buf, len);
  return *this;
} /* Format(const WCHAR *, ...) */

ASTR & ASTR::Format(const WCHAR *ctl, std::va_list list)
{ UA4     len;
  WCHAR   buf[4096];
  len = (UA4)vswprintf(buf, ctl, list);
  Set(buf, len);
  return *this;
} /* Format(const WCHAR *, std::va_list) */

void ASTR::Reserve(UA4 len)
{ if(m_Max<len)
  { SysFreeString(m_Str);
    if(m_Max==0) m_Max=64;
    while(m_Max<len) m_Max *= 2;
    m_Str = SysAllocStringLen(NULL, m_Max);
  }
} /* Reserve */

void ASTR::UpdateLength()
{ SetLength(m_Str==NULL ? 0 : (UA4)wcslen(m_Str));
} /* UpdateLength */

ASTR & ASTR::operator+=(WCHAR c)
{ if(m_Length==m_Max) Grow(m_Max+Max(m_Max/4, (UA4)64));
  m_Str[m_Length++]=c, m_Str[m_Length]=0;
  return *this;
} /* operator+=(WCHAR) */

ASTR ASTR::FormatNew(const WCHAR *ctl, ...)
{ std::va_list list;
  va_start(list, ctl);
  ASTR tmp;
  tmp.Format(ctl, list);
  va_end(list);
  return tmp;
} /* FormatNew */

void ASTR::Append(const WCHAR *s, UA4 len)
{ assert(s != NULL);
  if(len==(UA4)-1) len=(UA4)wcslen(s);
  UA4 nlen = m_Length+len;

  if(nlen > m_Max) Grow(nlen);
  std::wcscat(m_Str+m_Length, s);
  SetLength(nlen);
} /* Append */

void ASTR::Set(const WCHAR *s, UA4 len)
{ assert(s != NULL);
  if(len==(UA4)-1) len=(UA4)wcslen(s);
  if(len>m_Max) Reserve(len);
  std::memcpy(m_Str, s, (len+1)*sizeof(WCHAR));
  SetLength(len);
} /* Set */

void ASTR::Grow(UA4 len)
{ if(m_Max==0) m_Max=64;
  while(m_Max<len) m_Max *= 2;
  
  WCHAR *tmp = SysAllocStringLen(NULL, m_Max);
  std::memcpy(tmp, m_Str, (m_Length+1)*sizeof(WCHAR));
  SysFreeString(m_Str);
  m_Str = tmp;
} /* Grow */

const WCHAR ASTR::EmptyBSTR[1+(4/sizeof(WCHAR))] = {0, 0, 0};