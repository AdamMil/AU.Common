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
 
#ifndef BLOCKALLOC_H_INC
#define BLOCKALLOC_H_INC

#include <new>

template<typename T>
class CAllocBlock
{ public:
  CAllocBlock()  { m_Block=NULL; }
  CAllocBlock(UA4 size, bool bNew, bool bDel) { m_Block=NULL; Init(size, bNew, bDel); }
  ~CAllocBlock() { ::operator delete(m_Block); }

  void Init(UA4 size, bool bNew, bool bDelete);

  T * New();
  void Delete(T *);
  
  bool IsFull() const throw()         { return m_Head==NULL; }
  bool IsIn(const T *p) const throw() { return p>=m_Block && p<m_End; }
  
  private:
  void *m_Head;
  T    *m_Block, *m_End;
  bool  m_bNew,   m_bDel;
};

template<typename T>
class CBlockAlloc
{ public:
  CBlockAlloc() { m_nBlocks=m_maxBlocks=m_Size=0; }
  CBlockAlloc(UA4 size, bool bNew, bool bDelete, bool bMul=true) { m_Size=0; Init(size, bNew, bDelete, bMul); }
 ~CBlockAlloc();
 
  void Init(UA4 size, bool bNew, bool bDelete, bool bMul=true);

  T * New();
  void Delete(T *);
  
  private:
  CBlockAlloc(const CBlockAlloc<T> &);
  CBlockAlloc<T> & operator=(const CBlockAlloc<T> &);
  
  CAllocBlock<T> *m_Blocks;
  UA4   m_nBlocks, m_maxBlocks, m_Size;
  bool  m_bNew, m_bDel, m_bMul;
};

template<typename T> void CAllocBlock<T>::Init(UA4 size, bool bNew, bool bDelete)
{ assert(sizeof(T) >= sizeof(void*));
  assert(m_Block==NULL);
  if(size==0) size=256;
  m_Head  = ::operator new(sizeof(T)*size);
  m_Block = (T*)m_Head;
  m_End   = m_Block+size;
  m_bNew  = bNew, m_bDel = bDelete;

  UA4   i;
  for(i=0; i<size-1; i++) *(void**)(m_Block+i)=m_Block+i+1;
  *(void**)(m_Block+i) = NULL;
}

template<typename T> T * CAllocBlock<T>::New()
{ assert(m_Block != NULL);
  if(IsFull()) throw std::bad_alloc();
  T *ret = (T*)m_Head;
  m_Head = *(void**)m_Head;
  if(m_bNew) ::new (ret) T;
  return ret;
}

template<typename T> void CAllocBlock<T>::Delete(T *p)
{ if(!p) return;
  assert(m_Block != NULL);
  assert(IsIn(p));
  if(m_bDel) p->~T();
  *(void**)p = m_Head;
  m_Head = p;
}

template<typename T> void CBlockAlloc<T>::Init(UA4 start, bool bNew, bool bDelete, bool bMul)
{ assert(sizeof(T) >= sizeof(void*));
  assert(m_Size==0);
  if(start==0) start=256;
  m_Blocks  = (CAllocBlock<T>*)::operator new((m_maxBlocks=4)*sizeof(CAllocBlock<T>));
  m_nBlocks = 1;
  m_Size    = start;
  m_bNew    = bNew, m_bDel = bDelete, m_bMul = bMul;
  for(UA1 i=0; i<m_maxBlocks; i++) ::new (m_Blocks+i) CAllocBlock<T>;
  m_Blocks[0].Init(start, bNew, bDelete);
}

template<typename T> CBlockAlloc<T>::~CBlockAlloc()
{ for(UA4 i=0; i<m_nBlocks; i++) m_Blocks[i].~CAllocBlock();
  ::operator delete(m_Blocks);
}

template<typename T> T * CBlockAlloc<T>::New()
{ assert(m_Size != 0);
  UA4 i=m_nBlocks-1;
  do if(!m_Blocks[i].IsFull()) break; while(i--);
  if(i==UA4_MAX)
  { assert(m_Size != 0);
    i=m_nBlocks;
    if(m_nBlocks==m_maxBlocks)
    { CAllocBlock<T> *tmp = (CAllocBlock<T>*)::operator new((m_maxBlocks*=2)*sizeof(CAllocBlock<T>));
      memcpy(tmp, m_Blocks, m_nBlocks*sizeof(CAllocBlock<T>));
      ::operator delete(m_Blocks);
      m_Blocks = tmp;
      for(UA4 j=m_nBlocks; j<m_maxBlocks; j++) ::new (m_Blocks+j) CAllocBlock<T>;
    }
    m_Blocks[m_nBlocks++].Init(m_bMul ? (m_nBlocks ? m_Size*=2 : m_Size) : m_Size, m_bNew, m_bDel);
  }
  return m_Blocks[i].New();
}

template<typename T> void CBlockAlloc<T>::Delete(T *p)
{ UA4 i;
  if(!p) return;
  for(i=0; i<m_nBlocks; i++) if(m_Blocks[i].IsIn(p)) break;
  assert(i<m_nBlocks);
  m_Blocks[i].Delete(p);
}

#endif // BLOCKALLOC_H_INC
