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

#include "stdafx.h"
#include "Helpers.h"

/*** Globals ***/
VARIANT g_vMissing; // initialized in DllMain (Common.cpp)

/*** ASTRList implementation ***/
UA4 ASTRList::Add(const WCHAR *s, UA4 len)
{ UA4   ret  = (UA4)m_Vec.size();
  ASTR *pstr = new ASTR;
  pstr->Set(s, len);
  m_Vec.push_back(pstr);
  return ret;
} /* Add */

UA4 ASTRList::Add(ASTR &s)
{ UA4   ret  = (UA4)m_Vec.size();
  ASTR *pstr = new ASTR(s);
  m_Vec.push_back(pstr);
  return ret;
} /* Add */

void ASTRList::Clear()
{ UA4 i, len = Length();
  for(i=0; i<len; i++) delete m_Vec[i];
  m_Vec.clear();
} /* Clear */

ASTR ASTRList::Join(const WCHAR *s) const
{ ASTR ret;
  UA4  i, len = Length(), total, slen;
  if(len == 0) return ret;
  slen = (UA4)wcslen(s), total = slen*(len-1);
  for(i=0; i<len; i++) total += m_Vec[i]->Length();
  ret.Reserve(total);
  for(i=0; i<len; i++) ret += *m_Vec[i];
  return ret;
} /* Join */

/*** ASTR implementation ***/
ASTR & ASTR::Attach(BSTR str)
{ if(m_Str != EmptyBSTR) SysFreeString(m_Str);
  if(str)
  { m_Length = m_Max = (UA4)SysStringLen(str);
    m_Str    = str;
  }
  else Init();
  return *this;
} /* Attach */

BSTR ASTR::Detach()
{ BSTR ret=m_Str;
  Init();
  if(ret==EmptyBSTR) ret = SysAllocString(m_Str);
  return ret;
} /* Detach */

void ASTR::Clear()
{ if(m_Str != EmptyBSTR) SysFreeString(m_Str);
  Init();
} /* Clear */

ASTR & ASTR::Format(const WCHAR *ctl, ...)
{ SNP va_list list;
  UA4   len;
  WCHAR buf[4096];
  va_start(list, ctl);
  len = (UA4)vswprintf(buf, ctl, list);
  va_end(list);
  Set(buf, len);
  return *this;
} /* Format(const WCHAR *, ...) */

ASTR & ASTR::Format(const WCHAR *ctl, SNP va_list list)
{ UA4     len;
  WCHAR   buf[4096];
  len = (UA4)vswprintf(buf, ctl, list);
  Set(buf, len);
  return *this;
} /* Format(const WCHAR *, std::va_list) */

ASTR & ASTR::Replace(const WCHAR *find, const WCHAR *rep)
{ assert(find && rep);
  UA4 flen=(UA4)wcslen(find), rlen=(UA4)wcslen(rep), rbytes=rlen*sizeof(WCHAR);
  if(flen==0) return *this;
  WCHAR *p=wcsstr(m_Str, find);
  if(!p) return *this;

  if(flen < rlen)
  { ASTR tmp;
    WCHAR *q=m_Str;
    do
    { tmp.Append(q, (UA4)(p-q));
      tmp.Append(rep, rlen);
    } while((p=wcsstr(q=p+flen, find)) != NULL);
    if(q-m_Str != m_Length) tmp.Append(q, m_Length-(q-m_Str));
    Attach(tmp.Detach());
  }
  else if(flen > rlen)
  { WCHAR *q;
    UA4    diff=flen-rlen, dbytes=diff*sizeof(WCHAR), nrem=0;
    do
    { nrem += diff;
      memcpy(p, rep, rbytes);
      q = wcsstr(p+=flen, find);
      memmove(p+rlen, p, (q==NULL ? m_Str+m_Length-p : q-p)*sizeof(WCHAR));
    } while((p=q) != NULL);
    m_Str[m_Length-nrem]=0;
    SetLength(m_Length-nrem);
  }
  else do memcpy(p, rep, rbytes); while((p=wcsstr(p+=flen, find)) != NULL);
  return *this;
} /* Replace */

ASTR ASTR::Substring(IA4 start, IA4 end) const
{ ASTR ret;
  if(start<0) start += m_Length; // negative indexes become index from end
  if(end<0)   end   += m_Length;

  UA4 s=(UA4)start, e=(UA4)end;
  if(s > e || s >= m_Length) return ret;
  if(e >= m_Length) e = m_Length-1;

  ret.Set(m_Str+s, e-s+1);
  return ret;
} /* Substring */

ASTRList ASTR::Split(const WCHAR *sep) const
{ assert(sep);
  ASTRList list;
  if(!sep[0])
  { WCHAR c;
    for(UA4 i=0; i<m_Length; i++)
    { c=m_Str[i];
      list.Add(&c, 1);
    }
  }
  else
  { ASTR   tmp;
    WCHAR *p, *q=m_Str;
    UA4    slen=(UA4)wcslen(sep);
    while(p=wcsstr(q, sep))
    { list.Add(tmp.Set(q, (p-q)));
      q+=slen;
    }
    if(tmp.Set(q, (p-q)).Length()) tmp.Add(tmp);
  }
  return list;
} /* Split */

ASTR ASTR::ToLowerCase() const
{ return ASTR(*this).LowerCase();
} /* ToLowerCase */

ASTR ASTR::ToUpperCase() const
{ return ASTR(*this).UpperCase();
} /* ToUpperCase */

ASTR & ASTR::SetCStr(const char *s)
{ UA4 len = (UA4)strlen(s);
  if(len>m_Max) Grow(len);
  MultiByteToWideChar(CP_ACP, 0, s, (int)len, m_Str, (int)m_Max);
  m_Str[len]=0;
  return *this;
} /* SetCStr */

void ASTR::Reserve(UA4 len)
{ if(m_Max<len)
  { if(m_Str != EmptyBSTR) SysFreeString(m_Str);
    if(m_Max==0) m_Max=64;
    while(m_Max<len) m_Max *= 2;
    m_Str = SysAllocStringLen(NULL, m_Max);
    m_Str[0] = 0;
  }
  SetLength(0);
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
{ SNP va_list list;
  va_start(list, ctl);
  ASTR tmp;
  tmp.Format(ctl, list);
  va_end(list);
  return tmp;
} /* FormatNew */

ASTR & ASTR::Append(const WCHAR *s, UA4 len)
{ assert(s != NULL);
  if(len==(UA4)-1) len=(UA4)wcslen(s);
  UA4 nlen = m_Length+len;

  if(nlen > m_Max) Grow(nlen);
  SNP wcscat(m_Str+m_Length, s);
  SetLength(nlen);
  return *this;
} /* Append */

ASTR & ASTR::Set(const WCHAR *s, UA4 len)
{ assert(s != NULL);
  if(len==(UA4)-1) len=(UA4)wcslen(s);
  if(len>m_Max) Reserve(len);
  SNP memcpy(m_Str, s, len*sizeof(WCHAR));
  m_Str[len]=0;
  SetLength(len);
  return *this;
} /* Set(const WCHAR *, UA4) */

ASTR & ASTR::Set(WCHAR c)
{ if(m_Max<1) Reserve(1);
  m_Str[0] = c;
  m_Str[1] = 0;
  SetLength(1);
  return *this;
} /* Set(WCHAR) */

void ASTR::Grow(UA4 len)
{ if(m_Max==0) m_Max=64;
  while(m_Max<len) m_Max *= 2;
  
  WCHAR *tmp = SysAllocStringLen(NULL, m_Max);
  SNP memcpy(tmp, m_Str, (m_Length+1)*sizeof(WCHAR));
  if(m_Str != EmptyBSTR) SysFreeString(m_Str);
  m_Str = tmp;
} /* Grow */

const WCHAR ASTR::EmptyBSTR[1+(4/sizeof(WCHAR))] = {0, 0, 0};

/*** AVAR implementation ***/
void AVAR::Attach(const VARIANT &ov)
{ Clear();
  v=ov;
} /* Attach */

VARIANT AVAR::Detach()
{ VARIANT r=v;
  Clear();
  return r;
} /* Detach */

HRESULT AVAR::Copy(AVAR &out) const
{ if(&out.v==&v) return S_OK;
  if(!out.IsEmpty()) out.Clear();
  return VariantCopy(&out.v, const_cast<VARIANT*>(&v));
} /* Copy */

HRESULT AVAR::CopyTo(VARIANT *pv) const
{ assert(pv && pv->vt==VT_EMPTY);
  if(pv==&v) return S_OK;
  return VariantCopy(pv, const_cast<VARIANT*>(&v));
} /* CopyTo */

HRESULT AVAR::ChangeType(VARTYPE type)
{ if(type==v.vt) return S_OK;
  if(v.vt & VT_ARRAY)
  { VARIANT tmp;
    HRESULT hRet;
    VariantClear(&tmp);
    IFS(VariantChangeType(&tmp, &v, 0, type)) Attach(tmp);
    return hRet;
  }
  else return VariantChangeType(&v, &v, 0, type);
} /* ChangeType */

HRESULT AVAR::ChangeCopy(VARTYPE type, AVAR &out) const
{ if(type==v.vt) return Copy(out);
  HRESULT hRet;
  VARIANT copy, tmp;
  
  IFS(VariantCopy(&copy, const_cast<VARIANT*>(&v)))
  { IFS(VariantChangeType(&tmp, &copy, 0, type)) out.Attach(tmp);
    VariantClear(&copy);
  }
  return hRet;
} /* ChangeCopy */

int AVAR::Cmp(const VARIANT &rhs) const
{ if(v.vt==VT_NULL) return rhs.vt==VT_NULL ? 0 : -1;
  else if(rhs.vt==VT_NULL) return 1;
  HRESULT hr = VarCmp(const_cast<VARIANT*>(&v), const_cast<VARIANT*>(&rhs), 0, 0);
  return (int)hr-1;
} /* Cmp */

AVAR & AVAR::operator=(const VARIANT &rhs)
{ if(&rhs != &v)
  { Clear();
    VariantCopy(&v, const_cast<VARIANT*>(&rhs));
  }
  return *this;
} /* operator= */

/*** AXMLNode implementation ***/
AXMLNode AXMLNode::FirstChild() const
{ assert(p != NULL);
  IXMLDOMNode *pnode=NULL;
  p->get_firstChild(&pnode);
  return AXMLNode(pnode);
} /* FirstChild */

AXMLNode AXMLNode::LastChild() const
{ assert(p != NULL);
  IXMLDOMNode *pnode=NULL;
  p->get_lastChild(&pnode);
  return AXMLNode(pnode);
} /* LastChild */

AXMLNode AXMLNode::NextSibling() const
{ assert(p != NULL);
  IXMLDOMNode *pnode=NULL;
  p->get_nextSibling(&pnode);
  return AXMLNode(pnode);
} /* NextSibling */

AXMLNode AXMLNode::PrevSibling() const
{ assert(p != NULL);
  IXMLDOMNode *pnode=NULL;
  p->get_previousSibling(&pnode);
  return AXMLNode(pnode);
} /* PrevSibling */

AXMLNode AXMLNode::SelectSingleNode(BSTR path) const
{ assert(p != NULL);
  IXMLDOMNode *pnode=NULL;
  p->selectSingleNode(path, &pnode);
  return AXMLNode(pnode);
} /* SelectSingleNode */

AXMLNodeList AXMLNode::ChildNodes() const
{ assert(p != NULL);
  IXMLDOMNodeList *plist=NULL;
  p->get_childNodes(&plist);
  return AXMLNodeList(plist);
} /* ChildNodes */

AXMLNodeList AXMLNode::SelectNodes(BSTR path) const
{ assert(p != NULL);
  IXMLDOMNodeList *plist=NULL;
  p->selectNodes(path, &plist);
  return AXMLNodeList(plist);
} /* SelectNodes */

ASTR AXMLNode::NodeName() const
{ assert(p != NULL);
  ASTR ret;
  BSTR str=NULL;
  p->get_nodeName(&str);
  return ret.Attach(str);
} /* NodeName */

ASTR AXMLNode::Attr(BSTR attr) const
{ assert(p != NULL);
  AComPtr<IXMLDOMNamedNodeMap> attrs;
  AXMLNode node;
  HRESULT  hRet;
  IFS(p->get_attributes(&attrs))
    IFS(attrs->getNamedItem(attr, &node))
      return node.Text();
  return ASTR();
} /* Attr */
  
ASTR AXMLNode::Text() const
{ assert(p != NULL);
  ASTR ret;
  BSTR str=NULL;
  p->get_text(&str);
  return ret.Attach(str);
} /* Text */

AVAR AXMLNode::Value() const
{ assert(p != NULL);
  AVAR ret;
  p->get_nodeValue(&ret);
  return ret;
} /* Value */

/*** AXMLNodeList implementation ***/
UA4 AXMLNodeList::Length()
{ assert(p != NULL);
  long len;
  p->get_length(&len);
  return (UA4)len;
} /* Length */

AXMLNode AXMLNodeList::Item(UA4 index)
{ assert(p != NULL);
  AXMLNode node;
  p->get_item((long)index, &node);
  return node;
} /* Item */

void AXMLNodeList::Reset()
{ assert(p != NULL);
  p->reset();
} /* Reset */

AXMLNode AXMLNodeList::Next()
{ assert(p != NULL);
  AXMLNode node;
  p->nextNode(&node);
  return node;
} /* Next */

