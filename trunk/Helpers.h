/* 
   This file is part of the AU.Common library, a set of ActiveX
   controls to aid in COM and Web development.
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

#ifndef AU_COMMON_HELPERS_H
#define AU_COMMON_HELPERS_H

#include <cstdarg>
#include <vector>

#define SAFEARRAY(t) SAFEARRAY *
#define CHKRET(c)    { if(FAILED(hRet=(c))) return hRet; }
#define CHKRT2(c, l) { if(FAILED(hRet=(c))) goto l;      }
#define IFF(c)       if(FAILED(hRet=(c)))
#define IFS(c)       if(SUCCEEDED(hRet=(c)))

#define CREATE(c,i,v)      (v).CoCreateInstance(CLSID_##c, IID_##i)
#define CREATENS(ns,c,i,v) (v).CoCreateInstance(ns::CLSID_##c, ns::IID_##i)
#define IQUERY(v,i,d)      (v).QueryInterface(IID_##i, d)
#define DQUERY(v,i,d)      (v)->QueryInterface(IID_##i, (void**)&(d))

#define VBTRUE  (-1)
#define VBFALSE 0

#if _MSC_VER >= 1300
  #define SNP std::
#else
  #define SNP
#endif

typedef signed  char     I1;
typedef unsigned char    U1;
typedef short            I2;
typedef unsigned short   U2;
typedef int              I4;
typedef unsigned         U4;
typedef __int64          I8;
typedef unsigned __int64 U8;
typedef int              IA1; // at least 1 byte, but fastest size for machine
typedef unsigned         UA1; // ditto
typedef int              IA2; // at least 2 bytes, but fastest size for machine
typedef unsigned         UA2; // ditto
typedef int              IA4; // at least 4 bytes, but fastest size for machine
typedef unsigned         UA4; // ditto
typedef int              IMI; // machine int
typedef unsigned         UMI;

template<typename T> T Min(T a, T b) { return a<b ? a : b; }
template<typename T> T Max(T a, T b) { return a<b ? b : a; }
template<typename T> void Swap(T &a, T &b) { T t=a; a=b, b=t; }

/* ASTRList */
class ASTR;
class ASTRList
{ public:
 ~ASTRList() { Clear(); }
 
  UA4  Add(ASTR &);
  void Clear();
  UA4  Length() const { return (UA4)m_Vec.size(); }
  ASTR Join(const WCHAR *) const;

  ASTR & operator[](int i) { return *m_Vec[i]; }
  const ASTR & operator[](int i) const { return *m_Vec[i]; }
  
  private:
  std::vector<ASTR *> m_Vec;
};

/* ASTR */
class ASTR
{ public:
  ASTR() { Init(); }
  ASTR(const WCHAR *s) { Init(); Set(s); }
  ASTR(const ASTR &s)  { Init(); Set(s); }
  ASTR(UA4 len) { Init(); Reserve(len);  }
 ~ASTR() { SysFreeString(m_Str);         }

  operator WCHAR *()             { return m_Str; }
  operator const WCHAR *() const { return m_Str; }

  ASTR &  Attach(BSTR str);
  BSTR    Detach();
  void    Clear();

  ASTR &   Format(const WCHAR *ctl, ...);
  ASTR &   Format(const WCHAR *ctl, SNP va_list);
  ASTR &   Replace(const WCHAR *find, const WCHAR *rep);
  ASTR     Substr(IA4 start)          const { return Substring(start, (IA4)m_Length-1);  }
  ASTR     Substr(IA4 start, UA4 len) const { return Substring(start, start+(IA4)len-1); }
  ASTR     Substring(IA4 start, IA4 end) const;
  ASTRList Split(const WCHAR *) const;
  ASTR     ToLowerCase() const;
  ASTR     ToUpperCase() const;
  ASTR &   LowerCase() { wcslwr(m_Str); return *this; }
  ASTR &   UpperCase() { wcsupr(m_Str); return *this; }

  ASTR &   SetCStr(const char *);

  UA4     Length() const { return m_Length; }
  void    Reserve(UA4);
  void    SetLength(UA4 len) { m_Length=len, *((U4*)m_Str-1)=(U4)(len<<1); }
  void    UpdateLength();

  BSTR    ToBSTR() const { return SysAllocString(m_Str); }

  bool operator==(const WCHAR *rhs)  const { return wcscmp(m_Str, rhs)==0; }
  bool operator!=(const WCHAR *rhs)  const { return wcscmp(m_Str, rhs)!=0; }
  bool operator<(const WCHAR *rhs)   const { return wcscmp(m_Str, rhs)<0;  }
  bool operator>(const WCHAR *rhs)   const { return wcscmp(m_Str, rhs)>0;  }
  bool operator<=(const WCHAR *rhs)  const { return wcscmp(m_Str, rhs)<=0; }
  bool operator>=(const WCHAR *rhs)  const { return wcscmp(m_Str, rhs)>=0; }

  bool operator!() const { return m_Str[0] == 0; }
  bool IsEmpty()   const { return m_Str[0] == 0; }
  
  ASTR & operator=(const WCHAR *rhs) { Set(rhs); return *this; }
  ASTR & operator=(const ASTR  &rhs) { Set(rhs, rhs.Length()); return *this; }

  ASTR operator+(WCHAR w)        const { return ASTR(*this)+=w; }
  ASTR operator+(const WCHAR *s) const { return ASTR(*this)+=s; }
  ASTR operator+(const ASTR  &s) const { return ASTR(*this)+=s; }

  ASTR & operator+=(WCHAR);
  ASTR & operator+=(const WCHAR *s) { Append(s); return *this; }
  ASTR & operator+=(const ASTR  &s) { Append(s, s.Length()); return *this; }
  
  static ASTR  FormatNew(const WCHAR *ctl, ...);

  static const WCHAR EmptyBSTR[1+(4/sizeof(WCHAR))];

  private:
  WCHAR * m_Str;
  UA4     m_Length, m_Max;
  
  void Init() { m_Str=(WCHAR *)EmptyBSTR, m_Length=m_Max=0; }
  void Append(const WCHAR *, UA4 len=(UA4)-1);
  void Set   (const WCHAR *, UA4 len=(UA4)-1);
  void Grow(UA4);

  friend bool operator==(const WCHAR *a, const ASTR &b);
  friend bool operator!=(const WCHAR *a, const ASTR &b);
  friend bool operator<(const WCHAR *a, const ASTR &b);
  friend bool operator>(const WCHAR *a, const ASTR &b);
  friend bool operator<=(const WCHAR *a, const ASTR &b);
  friend bool operator>=(const WCHAR *a, const ASTR &b);
  friend ASTR operator+(WCHAR c, const ASTR &s);
  friend ASTR operator+(const WCHAR *a, const ASTR &b);
};

inline bool operator==(const WCHAR *a, const ASTR &b) { assert(a != NULL); return wcscmp(a, b)==0;  }
inline bool operator!=(const WCHAR *a, const ASTR &b) { assert(a != NULL); return wcscmp(a, b)!=0;  }
inline bool operator<(const WCHAR *a, const ASTR &b)  { assert(a != NULL); return wcscmp(a, b)<0;   }
inline bool operator>(const WCHAR *a, const ASTR &b)  { assert(a != NULL); return wcscmp(a, b)>0;   }
inline bool operator<=(const WCHAR *a, const ASTR &b) { assert(a != NULL); return wcscmp(a, b)<=0;  }
inline bool operator>=(const WCHAR *a, const ASTR &b) { assert(a != NULL); return wcscmp(a, b)>=0;  }
inline ASTR operator+(WCHAR c, const ASTR &s)         { WCHAR buf[2] = {c, 0}; return ASTR(buf)+=s; }
inline ASTR operator+(const WCHAR *a, const ASTR &b)  { assert(a != NULL); return ASTR(a)+=b;       }

/* AComPtr */
template<typename T>
class _NoAddRefReleaseOnAComPtr : public T
{ private:
		STDMETHODIMP_(ULONG) AddRef()  { };
		STDMETHODIMP_(ULONG) Release() { };
};

template<typename T>
class AComPtr
{ public:
  AComPtr()      { p=NULL; }
  AComPtr(T *op) { if ((p=op) != NULL) p->AddRef(); }
  AComPtr(AComPtr<T> &op) { if ((p=op.p) != NULL) p->AddRef(); }
 ~AComPtr()      { if(p) p->Release(); }

  operator T*()   const { return (T*)p; }
  T & operator*() const { assert(p!=NULL); return *p; }

  _NoAddRefReleaseOnAComPtr<T> * operator->() const { assert(p!=NULL); return (_NoAddRefReleaseOnAComPtr<T>*)p; }
  T ** operator&() { assert(p==NULL); return &p;  }

  bool operator !()       const { return p==NULL; }
  bool operator <(T *rhs) const { return p<rhs;   }
  bool operator==(T *rhs) const { return p==rhs;  }

  AComPtr<T>& operator=(T *lp)          { return Set(lp);   }
  AComPtr<T>& operator=(AComPtr<T> &lp) { return Set(lp.p); }

  bool IsEqualObject(IUnknown *op)
  { if(p==NULL && pOther==NULL) return true;
    if(p==NULL || pOther==NULL) return false;
    AComPtr<IUnknown> p1, p2;
    p ->QueryInterface(IID_IUnknown, (void**)&p1);
    op->QueryInterface(IID_IUnknown, (void**)&p2);
    return p1==p2;
  }

  void Release()
  { if(p)
    { p->Release();
      p=NULL;
    }
  }

  void Attach(T *p2)  { if(p) p->Release(); p=p2;   }
  T *  Detach()       { T* pt=p; p=NULL; return pt; }

  template<typename Q>
  void CopyTo(Q **o)
  { 
  #ifndef NDEBUG
    Q *t = p;
  #endif
    assert(o != NULL);
    if(*o=p) p->AddRef();
  }

  HRESULT CoCreateInstance(REFCLSID rclsid, REFIID riid, LPUNKNOWN pUnkOuter=NULL, DWORD dwClsContext=CLSCTX_INPROC_SERVER)
  { assert(p == NULL);
    return ::CoCreateInstance(rclsid, pUnkOuter, dwClsContext, riid, (void**)&p);
  }

  template<typename Q>
  HRESULT QueryInterface(REFIID riid, Q **dp) const
  { assert(dp != NULL);
    *dp = NULL;
    return p->QueryInterface(riid, (void**)dp);
  }

  T *p;

  protected:
  AComPtr<T> & Set(T *op)
  { if(op) op->AddRef();
    if(p)  p->Release();
    p=op;
    return *this;
  }
};

/* AVAR */
class AVAR
{ public:
  AVAR()                  { VariantInit(&v); }
  AVAR(const VARIANT &ov) { VariantInit(&v); VariantCopy(&v, const_cast<VARIANT*>(&ov));   }
  AVAR(const AVAR &ov)    { VariantInit(&v); VariantCopy(&v, const_cast<VARIANT*>(&ov.v)); }
 ~AVAR() { Clear(); }
  
  void    Clear() { VariantClear(&v); }
  bool    IsEmpty() const { return v.vt==VT_EMPTY; }
  VARTYPE Type()    const { return v.vt;           }

  void    Attach(const VARIANT &ov);
  VARIANT Detach();
  HRESULT Copy(AVAR &out) const;
  HRESULT CopyTo(VARIANT *pv) const;
  
  HRESULT ChangeType(VARTYPE type);
  HRESULT ChangeCopy(VARTYPE type, AVAR &out) const;
  BSTR    ToBSTR() const { assert(v.vt==VT_BSTR); return v.bstrVal; }
  
  void    SetI4(I4 i) { Clear(); v.vt=VT_I4, v.intVal=i; }

  int Cmp(const VARIANT &rhs) const;

  VARIANT * operator&() { assert(v.vt==VT_EMPTY); return &v; }
  const VARIANT * operator&() const { assert(v.vt==VT_EMPTY); return &v; }
  operator VARIANT &()  { return v; }

  AVAR & AVAR::operator=(const VARIANT &rhs);

  bool operator==(const VARIANT &rhs) const { return Cmp(rhs)==0; }
  bool operator!=(const VARIANT &rhs) const { return Cmp(rhs)!=0; }
  bool operator< (const VARIANT &rhs) const { return Cmp(rhs)<0;  }
  bool operator> (const VARIANT &rhs) const { return Cmp(rhs)>0;  }
  bool operator<=(const VARIANT &rhs) const { return Cmp(rhs)<=0; }
  bool operator>=(const VARIANT &rhs) const { return Cmp(rhs)>=0; }
  
  VARIANT v;
};
bool operator==(const VARIANT &lhs, const AVAR &rhs) { return rhs.Cmp(lhs)==0; }
bool operator!=(const VARIANT &lhs, const AVAR &rhs) { return rhs.Cmp(lhs)!=0; }
bool operator<(const VARIANT &lhs,  const AVAR &rhs) { return rhs.Cmp(lhs)>=0; }
bool operator>(const VARIANT &lhs,  const AVAR &rhs) { return rhs.Cmp(lhs)<=0; }
bool operator<=(const VARIANT &lhs, const AVAR &rhs) { return rhs.Cmp(lhs)>0;  }
bool operator>=(const VARIANT &lhs, const AVAR &rhs) { return rhs.Cmp(lhs)<0;  }

/* AXMLNode */
class AXMLNode : public AComPtr<IXMLDOMNode>
{ protected:
  typedef IXMLDOMNode Node;
  typedef AComPtr<IXMLDOMNode> Base;

  public:
  AXMLNode() { }
  AXMLNode(Node *op) : Base(op) { }
  AXMLNode(Base &op) : Base(op) { }

  AXMLNode & operator=(Node *op) { Base::operator=(op); return *this; }

  AXMLNode FirstChild() const;
  AXMLNode LastChild() const;
  AXMLNode NextSibling() const;
  AXMLNode PrevSibling() const;
  AXMLNode SelectSingleNode(BSTR path) const;

  IXMLDOMNodeList * ChildNodes() const;
  IXMLDOMNodeList * SelectNodes(BSTR path) const;
  
  ASTR     NodeName() const;
  ASTR     Attr(BSTR attr) const;
  ASTR     Text() const;
  AVAR     Value() const;
};

/* AXMLNodeList */
class AXMLNodeList : public AComPtr<IXMLDOMNodeList>
{ protected:
  typedef IXMLDOMNodeList List;
  typedef AComPtr<IXMLDOMNodeList> Base;

  public:
  AXMLNodeList() { }
  AXMLNodeList(List *op) : Base(op) { }
  AXMLNodeList(Base &op) : Base(op) { }

  AXMLNodeList & operator=(List *op) { Base::operator=(op); return *this; }

  UA4 Length();
  AXMLNode Item(UA4 index);
  
  void Reset();
  AXMLNode Next();
};

/* ACritSec */
class ACritSec
{ public:
  ACritSec() { InitializeCriticalSection(&m_CS); }
 ~ACritSec() { DeleteCriticalSection(&m_CS);     }
 
  void Lock()   { EnterCriticalSection(&m_CS); }
  void Unlock() { LeaveCriticalSection(&m_CS); }

  private:
  CRITICAL_SECTION m_CS;  
};

/* AAutoLock */
class AAutoLock
{ public:
  AAutoLock(ACritSec &cs, bool bLocked=false) : m_CS(cs) { m_bLocked=bLocked; }
 ~AAutoLock() { if(m_bLocked) Unlock(); }

  void Lock()   { assert(!m_bLocked); m_CS.Lock();   m_bLocked=true;  }
  void Unlock() { assert(m_bLocked);  m_CS.Unlock(); m_bLocked=false; }

  private:
  ACritSec &m_CS;
  bool      m_bLocked;
};

// defined in Helpers.cpp
extern VARIANT g_vMissing;

#endif
