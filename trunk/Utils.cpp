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
#include "Utils.h"
#include "Config.h"

/*** statics ***/
AComPtr<IConfig> s_Config;
HINSTANCE        s_DllHInst;

class SHA1Digest
{ public:
    SHA1Digest();
    HRESULT Hash(void *in, UA4 inlen, U1 *out);
    U4  m_Hash[5], m_Length;
    UA1 m_Pos;
    U1  m_Buf [64];

  private:
    void Crunch();
    void Pad();
    inline U4 Shift(U4 bits, U4 word) { return (word<<bits)|(word>>(32-bits)); }
};

static const U1 *s_HexTable=(const U1*)"0123456789ABCDEF";
static const U1  s_B64Enc[64] = 
{ 65,  66,  67,  68,  69,  70,  71,  72,  73,  74,  75,  76,  77,  78,  79,  80,
  81,  82,  83,  84,  85,  86,  87,  88,  89,  90,  97,  98,  99,  100, 101, 102,
  103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118,
  119, 120, 121, 122, 48,  49,  50,  51,  52,  53,  54,  55,  56,  57,  43,  47
};

static const U1  s_B64Dec[256] =
{ 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,
  255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,
  255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 62,  255, 255, 255, 63,
  52,  53,  54,  55,  56,  57,  58,  59,  60,  61,  255, 255, 255, 64,  255, 255,
  255, 0,   1,   2,   3,   4,   5,   6,   7,   8,   9,   10,  11,  12,  13,  14,
  15,  16,  17,  18,  19,  20,  21,  22,  23,  24,  25,  255, 255, 255, 255, 255,
  255, 26,  27,  28,  29,  30,  31,  32,  33,  34,  35,  36,  37,  38,  39,  40,
  41,  42,  43,  44,  45,  46,  47,  48,  49,  50,  51,  255, 255, 255, 255, 255,
  255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,
  255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,
  255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,
  255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,
  255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,
  255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,
  255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,
  255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255
};

/*** rs fields ***/
HRESULT g_GetField(ADORecordset *rs, UA4 nField, VARIANT *out)
{ VARIANT var;
  var.vt=VT_I4, var.intVal=(int)nField;
  return g_GetField(rs, var, out);
} /* g_GetField(ADORecordset *, UA4, VARIANT *) */

HRESULT g_GetField(ADORecordset *rs, VARIANT vField, VARIANT *out)
{ assert(out != NULL);
  if(out==NULL) return E_POINTER;
  AComPtr<ADOFields> fields;
  AComPtr<ADOField>  field;
  HRESULT hRet;
  CHKRET(rs->get_Fields(&fields));
  CHKRET(fields->get_Item(vField, &field));
  CHKRET(field->get_Value(out));
  return hRet;
} /* g_GetField(ADORecordset *, VARIANT, VARIANT *) */

/*** config ***/
HRESULT g_Config(BSTR sKey, BSTR sType, BSTR sSection, AVAR &out)
{ if(!s_Config)
  { HRESULT hRet;
    CHKRET(g_LoadConfig());
  }
  out.Clear();
  return s_Config->get_Item(sKey, sType, sSection, &out);
} /* g_Config */

void g_InitConfig(HINSTANCE hInst)
{ s_DllHInst = hInst;
} /* g_InitConfig */

HRESULT g_LoadConfig()
{ HRESULT hRet;
  char    buf[MAX_PATH+64], *p;
  if(!GetModuleFileName(s_DllHInst, buf, MAX_PATH)) return E_FAIL;
  p=strrchr(buf, L'\\');
  if(!p++) return E_FAIL;
  strcpy(p, "AU.Common.xml");
  if(!s_Config) CHKRET(CREATE(Config, IConfig, s_Config));
  IFF(s_Config->OpenFile(ASTR::FromCStr(buf)))
    hRet=s_Config->OpenXML(ASTR(L"<config/>"));
  return hRet;
} /* g_LoadConfig */

void g_UnloadConfig()
{ // have to wrap this because if the xml dll has already been unloaded,
  // there will be an access violation. the approach is not good, but
  // I don't believe there's a way to control the order of COM server
  // loading/unloading or to detect the application end before
  // CoUnitialize is called
  try { s_Config.Release(); } catch(...) { s_Config.p=NULL; }
} /* g_UnloadConfig */

HRESULT g_StringToBin(BSTR sStr, bool b8bit, AVAR &out)
{ UA4  i, len = SysStringByteLen(sStr);
  if(len==0)
  { out.Clear();
    out.v.vt=VT_NULL;
    return S_OK;
  }

  SAFEARRAYBOUND bound = { b8bit ? len=(len+1)/2 : len, 0 };
  SAFEARRAY     *sa    = SafeArrayCreate(VT_UI1, 1, &bound);
  U1            *buf   = (U1*)sa->pvData;
  
  if(b8bit) for(i=0; i<len; i++) buf[i]=(U1)sStr[i];
  else memcpy(buf, sStr, len);
  out.Clear();
  out.v.vt = VT_ARRAY|VT_UI1, out.v.parray = sa;
  return S_OK;
} /* g_StringToBin */

HRESULT g_BinToString(VARIANT &vBin, bool b8bit, ASTR &out)
{ if(vBin.vt == VT_NULL)
  { out.Clear();
    return S_OK;
  }

  SAFEARRAY *sa = vBin.parray;
  BSTR       bstr;
  UA4        len;
  if((vBin.vt&VT_ARRAY)==0 || sa->cbElements != 1) return E_INVALIDARG;

  len = sa->rgsabound[0].cElements;
  if(b8bit)
  { U1 *buf = (U1*)sa->pvData;
    UA4 i;
    bstr = SysAllocStringLen(NULL, len);
    for(i=0; i<len; i++) bstr[i]=buf[i];
  }
  else
  { bstr = SysAllocStringByteLen(NULL, len);
    memcpy(bstr, sa->pvData, len);
  }
  out.Attach(bstr);
  return S_OK;
} /* g_BinToString */

HRESULT g_HexToBin(BSTR sHex, AVAR &out)
{ UA4  i, j, len = SysStringLen(sHex);
  if(len&1) return E_INVALIDARG;
  if(len==0)
  { out.Clear();
    out.v.vt=VT_NULL;
    return S_OK;
  }

  SAFEARRAYBOUND bound = { len/2, 0 };
  SAFEARRAY     *sa    = SafeArrayCreate(VT_UI1, 1, &bound);
  U1            *buf   = (U1*)sa->pvData, high, low;
  for(i=j=0; i<len; i+=2)
  { high=sHex[i]-L'0', low=sHex[i+1]-L'0';
    if(high>9) { high-=7; if(high>15) goto Err; }
    if(low>9)  { low -=7; if(low >15) goto Err; }
    buf[j++] = (high<<4)|low;
  }
  out.Clear();
  out.v.vt = VT_ARRAY|VT_UI1, out.v.parray = sa;
  return S_OK;

  Err:
  return E_INVALIDARG;
} /* g_HexToBin */

HRESULT g_BinToHex(VARIANT &vBin, bool b8bit, ASTR &out)
{ if(vBin.vt==VT_NULL)
  { out.Clear();
    return S_OK;
  }
  if(vBin.vt != VT_BSTR && ((vBin.vt&VT_ARRAY)==0 || vBin.parray->cbElements!=1)) return E_INVALIDARG;

  BSTR ret;
  UA4  i, j, len;
  U1   c;
  if(b8bit)
  { if(vBin.vt == VT_BSTR) len=SysStringLen(vBin.bstrVal);
    else len=vBin.parray->rgsabound[0].cElements;
  }
  else
  { if(vBin.vt != VT_BSTR) return E_INVALIDARG;
    len = SysStringByteLen(vBin.bstrVal);
  }

  ret = SysAllocStringLen(NULL, len*2);
  if(b8bit && vBin.vt == VT_BSTR)
  { WCHAR *buf = vBin.bstrVal;
    for(i=j=0; i<len; j+=2,i++)
    { c=(U1)buf[i];
      ret[j]=s_HexTable[c>>4], ret[j+1]=s_HexTable[c&0xF];
    }
  }
  else
  { U1 *buf = vBin.vt==VT_BSTR ? (U1*)vBin.bstrVal : (U1*)vBin.parray->pvData;
    for(i=j=0; i<len; j+=2,i++)
      c=buf[i], ret[j]=s_HexTable[c>>4], ret[j+1]=s_HexTable[c&0xF];
  }

  out.Attach(ret);
  return S_OK;
} /* g_BinToHex */

HRESULT g_EncodeB64(VARIANT &vIn, bool b8bit, ASTR &out)
{ U1  *buf;
  BSTR ret;
  UA4  i, j, len, obytes;
  AVAR var;
  UA1  left;
  U1   bin[3];

  if(vIn.vt != VT_BSTR && ((vIn.vt&VT_ARRAY)==0 || vIn.parray->cbElements!=1)) return E_INVALIDARG;
  if(vIn.vt==VT_BSTR)
  { if(b8bit)
    { g_StringToBin(vIn.bstrVal, true, var);
      buf = (U1*)var.v.parray->pvData, len = var.v.parray->rgsabound[0].cElements;
    }
    else buf = (U1*)vIn.bstrVal, len = SysStringByteLen(vIn.bstrVal);
  }
  else
  { if(!b8bit) return E_INVALIDARG;
    buf = (U1*)vIn.parray->pvData, len = vIn.parray->rgsabound[0].cElements;
  }
  
  left   = (U1)len%3;
  obytes = (left==0) ? len/3*4 : (len+(3-left))/3*4;
  ret    = SysAllocStringLen(NULL, obytes);
  
  for(i=j=0,len-=left; i<len; buf+=3,j+=4,i+=3)
  { bin[0]=buf[0], bin[1]=buf[1], bin[2]=buf[2];
    ret[j  ]=s_B64Enc[bin[0]>>2];
    ret[j+1]=s_B64Enc[((bin[0]&3)<<4) | (bin[1]>>4)];
    ret[j+2]=s_B64Enc[((bin[1]&0xF)<<2) | (bin[2]>>6)];
    ret[j+3]=s_B64Enc[bin[2]&0x3F];
  }

  if(left==1)
  { bin[0]=*buf;
    ret[j  ]=s_B64Enc[bin[0]>>2];
    ret[j+1]=s_B64Enc[(bin[0]&3)<<4];
    ret[j+2]=ret[j+3]=61;
  }
  else if(left==2)
  { bin[0]=buf[0], bin[1]=buf[1];
    ret[j  ]=s_B64Enc[bin[0]>>2];
    ret[j+1]=s_B64Enc[((bin[0]&3)<<4)|(bin[1]>>4)];
    ret[j+2]=s_B64Enc[(bin[1]&0xF)<<2];
    ret[j+3]=61;
  }

  out.Attach(ret);
  return S_OK;
} /* g_EncodeB64 */

HRESULT g_DecodeB64(VARIANT &vIn, bool b8bit, AVAR &out)
{ U1  bin[4];
  UA4 i, j, len, obytes=0;
  UA1 left=0;

  if(vIn.vt==VT_NULL)
  { out.v.vt=VT_NULL;
    return S_OK;
  }
  if(vIn.vt != VT_BSTR && ((vIn.vt&VT_ARRAY)==0 || vIn.parray->cbElements!=1)) return E_INVALIDARG;
  if(b8bit && vIn.vt==VT_BSTR)
  { WCHAR *buf=vIn.bstrVal;
    len = SysStringLen(vIn.bstrVal);
    if(len==0)
    { out.v.vt=VT_NULL;
      return S_OK;
    }
    if(buf[len-1]==61)
    { left++;
      if(buf[len-2]==61) left++;
    }

    SAFEARRAYBOUND bound = { len*4/3, 0 };
    SAFEARRAY     *sa    = SafeArrayCreate(VT_UI1, 1, &bound);
    U1            *ret   = (U1*)sa->pvData;
    
    for(i=j=0; i<len; buf+=4,j+=3,i+=4)
    { bin[0]=s_B64Dec[(U1)buf[0]], bin[1]=s_B64Dec[(U1)buf[1]];
      bin[2]=s_B64Dec[(U1)buf[2]], bin[3]=s_B64Dec[(U1)buf[3]];
      if(bin[3]==64 || bin[2]==64)
      { if(bin[2] != 64)
        { ret[j  ]=(bin[0]<<2) | ((bin[1]>>4)&3);
          ret[j+1]=((bin[1]&0xF)<<4) | ((bin[2]>>2)&0xF);
        }
        else ret[j]=(bin[0]<<2) | ((bin[1]>>4)&3);
        break;
      }
      ret[j  ]=(bin[0]<<2) | ((bin[1]>>4)&3);
      ret[j+1]=((bin[1]&0xF)<<4) | ((bin[2]>>2)&0xF);
      ret[j+2]=((bin[2]&3)<<6) | bin[3];
    }
    
    out.Clear();
    out.v.vt=VT_ARRAY|VT_UI1, out.v.parray=sa;
  }
  else if(!b8bit && vIn.vt != VT_BSTR) return E_INVALIDARG;
  else
  { U1 *buf;
    if(vIn.vt==VT_BSTR) buf=(U1*)vIn.bstrVal, len=SysStringByteLen(vIn.bstrVal);
    else buf=(U1*)vIn.parray->pvData, len=vIn.parray->rgsabound[0].cElements;
    if(len==0)
    { out.v.vt=VT_NULL;
      return S_OK;
    }
    if(buf[len-1]==61)
    { left++;
      if(buf[len-2]==61) left++;
    }

    SAFEARRAYBOUND bound = { len*4/3, 0 };
    SAFEARRAY     *sa    = SafeArrayCreate(VT_UI1, 1, &bound);
    U1            *ret   = (U1*)sa->pvData;

    for(i=j=0; i<len; buf+=4,j+=3,i+=4)
    { *(U4*)bin = *(U4*)buf;
      bin[0]=s_B64Dec[bin[0]], bin[1]=s_B64Dec[bin[1]];
      bin[2]=s_B64Dec[bin[2]], bin[3]=s_B64Dec[bin[3]];
      if(bin[3]==64 || bin[2]==64)
      { if(bin[2] != 64)
        { ret[j  ]=(bin[0]<<2) | ((bin[1]>>4)&3);
          ret[j+1]=((bin[1]&0xF)<<4) | ((bin[2]>>2)&0xF);
        }
        else ret[j]=(bin[0]<<2) | ((bin[1]>>4)&3);
        break;
      }
      ret[j  ]=(bin[0]<<2) | ((bin[1]>>4)&3);
      ret[j+1]=((bin[1]&0xF)<<4) | ((bin[2]>>2)&0xF);
      ret[j+2]=((bin[2]&3)<<6) | bin[3];
    }
    
    out.Clear();
    out.v.vt=VT_ARRAY|VT_UI1, out.v.parray=sa;
  }
  return S_OK;
} /* g_DecodeB64 */

HRESULT g_HashSHA1(VARIANT &vIn, AVAR &out)
{ void *buf;
  UA4   len;

  if(vIn.vt == VT_BSTR) buf=vIn.bstrVal, len=SysStringByteLen(vIn.bstrVal);
  else if((vIn.vt&VT_ARRAY)!=0 && vIn.parray->cbElements==1)
    buf=vIn.parray->pvData, len=vIn.parray->rgsabound[0].cElements;
  else if(vIn.vt==VT_NULL) len=0;
  else return E_INVALIDARG;
  
  SHA1Digest     hash;
  SAFEARRAYBOUND bound = { 20, 0 };
  SAFEARRAY     *sa    = SafeArrayCreate(VT_UI1, 1, &bound);
  HRESULT        hRet  = hash.Hash(buf, len, (U1*)sa->pvData);

  if(SUCCEEDED(hRet))
  { out.Clear();
    out.v.vt=VT_UI1|VT_ARRAY, out.v.parray=sa;
  }
  else SafeArrayDestroy(sa);
  return hRet;
} /* g_HashSHA1 */

HRESULT g_CheckSHA1(VARIANT &vIn, VARIANT &vHash, bool &ok)
{ if((vHash.vt&VT_ARRAY)==0 || vHash.parray->cbElements!=1 ||
     vHash.parray->rgsabound[0].cElements!=20)
    return E_INVALIDARG;

  AVAR    hash;
  HRESULT hRet;
  CHKRET(g_HashSHA1(vIn, hash));
  ok = memcmp(hash.v.parray->pvData, vHash.parray->pvData, 20)==0;
  return S_OK;
} /* g_CheckSHA1 */

/*** SHA1Digest implementation ***/
SHA1Digest::SHA1Digest()
{ m_Hash[0]=0x67452301, m_Hash[1]=0xEFCDAB89, m_Hash[2]=0x98BADCFE, m_Hash[3]=0x10325476, m_Hash[4]=0xC3D2E1F0;
  m_Length = m_Pos = 0;
} /* SHA1Digest */

HRESULT SHA1Digest::Hash(void *in, UA4 inlen, U1 *out)
{ if(in==NULL || out==NULL) return E_POINTER;
  UA4 i=0;
  UA1 bytes;

  if(inlen > 0)
  { m_Length=inlen<<3;
    do
    { bytes = Min(inlen, (UA4)64);
      memcpy(m_Buf, (U1*)in+i, m_Pos=bytes);
      if(bytes < 64) Pad();
      Crunch();
      i += bytes, inlen-=bytes;
    } while(inlen != 0);
    if(bytes==64)
    { Pad();
      Crunch();
    }
  }

  for(i=0; i<20; i++) out[i]=(U1)(m_Hash[i>>2] >> 8*(3-(i&3)));
  return S_OK;
}

void SHA1Digest::Crunch()
{ static const U4 consts[] = { 0x5A827999, 0x6ED9EBA1, 0x8F1BBCDC, 0xCA62C1D6 };
  UA2 i, j;
  U4  t, a, b, c, d, e, buf[80];

  for(i=j=0; i<16; j+=4,i++)
    buf[i] = (m_Buf[j]<<24)|(m_Buf[j+1]<<16)|(m_Buf[j+2]<<8)|m_Buf[j+3];
  for(i=16; i<80; i++) buf[i] = Shift(1, buf[i-3]^buf[i-8]^buf[i-14]^buf[i-16]);
  a=m_Hash[0], b=m_Hash[1], c=m_Hash[2], d=m_Hash[3], e=m_Hash[4];
  for(i=0; i<20; i++)
  { t = Shift(5, a) + ((b & c) | ((~b) & d)) + e + buf[i] + consts[0];
    e=d, d=c, c=Shift(30, b), b=a, a=t;
  }
  for(; i<40; i++)
  { t = Shift(5, a) + (b ^ c ^ d) + e + buf[i] + consts[1];
    e=d, d=c, c=Shift(30, b), b=a, a=t;
  }
  for(; i<60; i++)
  { t = Shift(5, a) + ((b & c) | (b & d) | (c & d)) + e + buf[i] + consts[2];
    e=d, d=c, c=Shift(30, b), b=a, a=t;
  }
  for(; i<80; i++)
  { t = Shift(5, a) + (b ^ c ^ d) + e + buf[i] + consts[3];
    e=d, d=c, c=Shift(30, b), b=a, a=t;
  }
  m_Hash[0]+=a, m_Hash[1]+=b, m_Hash[2]+=c, m_Hash[3]+=d, m_Hash[4]+=e;
  m_Pos=0;
} /* Crunch */

void SHA1Digest::Pad()
{ m_Buf[m_Pos++]=0x80;
  if(m_Pos>56)
  { memset(m_Buf+m_Pos, 0, 64-m_Pos);
    Crunch();
    memset(m_Buf, 0, m_Pos=56);
  }
  else memset(m_Buf+m_Pos, 0, 56-m_Pos);
  *(U4*)(m_Buf+56)=0;
  m_Buf[60]=(U1)(m_Length>>24), m_Buf[61]=(U1)(m_Length>>16);
  m_Buf[62]=(U1)(m_Length>>8),  m_Buf[63]=(U1)m_Length;
} /* Pad */
