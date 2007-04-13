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

#include "stdafx.h"
#include "Utils.h"
#include "Config.h"
#include <mtx.h>

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

HRESULT g_CreateDB(IDB **ppdb, bool share, const WCHAR *key, const WCHAR *sect)
{ assert(ppdb);
  AComPtr<IDBMan> dbman;
  AVAR            var;
  HRESULT         hRet = share ? g_SessVar(L"oDBMan", &var) : E_FAIL;

  if(FAILED(hRet) || var.IsEmpty())
  { if(share)
    { static AComPtr<IDBMan> man;
      if(!man) CHKRET(CREATE(DBMan, IDBMan, man));
      dbman = man;
    }
    else
    { AComPtr<IDB> db;
      CHKRET(CREATE(DB, IDB, db));
      if(key &&key [0]) db->put_ConnKey(ASTR(key));
      if(sect&&sect[0]) db->put_ConnSection(ASTR(sect));
      db.CopyTo(ppdb);
      return hRet;
    }
  }
  else
  { if(var.Type() != VT_DISPATCH) return E_ABORT;
    CHKRET(DQUERY(var.v.pdispVal, IDBMan, dbman));
  }
  hRet = dbman->CreateDB(ASTR(key), ASTR(sect), ppdb);
  return hRet;
} /* g_CreateDB */

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

/*** other db stuff ***/
HRESULT g_DBCheckType(VARIANT *pv, bool alter)
{ VARTYPE vt = pv->vt;
  if(vt & VT_ARRAY) return E_INVALIDARG;
  while((vt&VT_TYPEMASK)==VT_VARIANT)
  { if(alter)
    { *pv=*pv->pvarVal, vt=pv->vt;
      if(vt & VT_ARRAY) return E_INVALIDARG;
    }
    else return g_DBCheckType(pv->pvarVal, false);
  }
  if(vt==VT_DISPATCH || vt==VT_UNKNOWN || vt>VT_UINT)
  { if(alter) return VariantChangeType(pv, pv, 0, VT_BSTR);
    else return E_INVALIDARG;
  }
  return S_OK;
} /* g_DBCheckType */

/*** session/application access ***/
HRESULT g_SessVar(BSTR sKey, VARIANT *pvout, bool statics)
{ HRESULT hRet = g_ASPSessVar(sKey, pvout, statics);
  if(FAILED(hRet) || pvout->vt==VT_EMPTY) hRet = g_AUSessVar(sKey, pvout);
  return hRet;
} /* g_SessVar */

HRESULT g_AUSessVar(BSTR sKey, VARIANT *pvout)
{ assert(pvout);
  if(!pvout) return E_POINTER;

  AComPtr<ISessionMan> sm;
  AComPtr<IAUSession>  sess;
  AVAR    var;
  HRESULT hRet;
  
  pvout->vt=VT_EMPTY;
  CHKRET(g_ASPAppVar(ASTR(L"oSessionMan"), &var, true));
  if(var.Type() != VT_DISPATCH) return S_FALSE;
  CHKRET(DQUERY(var.v.pdispVal, ISessionMan, sm));
  CHKRET(sm->get_Current(&sess));
  if(wcschr(sKey, L'.')) hRet = sess->get_Item(sKey, pvout);  
  else hRet = sess->get_Item(ASTR(L'.').Append(sKey), pvout);
  return hRet;
} /* g_AUSessVar */

HRESULT g_ASPVar(BSTR sKey, VARIANT *pvout, bool statics, bool bSess)
{ assert(pvout);
  if(!pvout) return E_POINTER;
  static const ASTR props[2] = { L"Application", L"Session" };
  static const IID *iids[2]  = { &ASP::IID_IApplicationObject, &ASP::IID_ISessionObject };

  AComPtr<IObjectContext>          context;
  AComPtr<IGetContextProperties>   cp;
  AComPtr<ASP::ISessionObject>     sess;
  AComPtr<ASP::IApplicationObject> app;
  AComPtr<ASP::IVariantDictionary> dict;
  AVAR    var;
  HRESULT hRet = GetObjectContext(&context);
  if(FAILED(hRet) && hRet != CONTEXT_E_NOCONTEXT) return hRet;
  pvout->vt = VT_EMPTY;
  if(!context) return S_OK;
  
  CHKRET(IQUERY(context, IGetContextProperties, &cp));
  CHKRET(cp->GetProperty((BSTR)props[bSess].InnerStr(), &var));
  if(var.Type() != VT_DISPATCH) return E_ABORT;
  CHKRET(var.v.pdispVal->QueryInterface(*iids[bSess], bSess ? (void**)&sess : (void**)&app));
  if(sess)
  { if(statics) CHKRET(sess->get_StaticObjects(&dict))
    else CHKRET(sess->get_Contents(&dict));
  }
  else
  { if(statics) CHKRET(app->get_StaticObjects(&dict))
    else CHKRET(app->get_Contents(&dict));
  }

  var.Clear();
  var.v.vt=VT_BSTR, var.v.bstrVal=sKey;
  hRet = dict->get_Item(var, pvout);
  var.Detach();
  return hRet;
} /* g_ASPVar */

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
  // there will be an access violation. this approach is not good, but
  // I don't believe there's a way to control the order of COM server
  // loading/unloading or to detect the application end before
  // CoUnitialize is called
  try { s_Config.Release(); } catch(...) { s_Config.p=NULL; }
} /* g_UnloadConfig */

/*** safearrays ***/
void g_InitSA(SAFEARRAY *sa, VARIANT *data, U4 els)
{ assert(sa != NULL);
  sa->cbElements = sizeof(VARIANT);
  sa->cDims      = 1;
  sa->cLocks     = 0;
  sa->fFeatures  = 0;
  sa->pvData     = data;
  sa->rgsabound[0].cElements = els;
  sa->rgsabound[0].lLbound   = 0;
} /* g_InitSA */

/*** randomness ***/
void g_RandBuf(void *b, UA4 len)
{ U1 *buf = (U1*)b;
  for(UA4 i=0; i<len; i++) buf[i] = rand() & 0xFF;
}

void g_RandStr(WCHAR *buf, UA4 len)
{ for(UA4 i=0; i<len; i++) buf[i] = rand()%95+32;
}

/*** data conversion ***/
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
  obytes = (left==0 ? len/3*4 : (len+(3-left))/3*4) + len/76;
  ret    = SysAllocStringLen(NULL, obytes);
  
  for(i=j=obytes=0,len-=left; i<len; buf+=3,j+=4,i+=3)
  { bin[0]=buf[0], bin[1]=buf[1], bin[2]=buf[2];
    ret[j  ]=s_B64Enc[bin[0]>>2];
    ret[j+1]=s_B64Enc[((bin[0]&3)<<4) | (bin[1]>>4)];
    ret[j+2]=s_B64Enc[((bin[1]&0xF)<<2) | (bin[2]>>6)];
    ret[j+3]=s_B64Enc[bin[2]&0x3F];
    if(++obytes==19) 
    { ret[j++ + 4]='\n';
      obytes = 0;
    }
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
  { WCHAR *buf=vIn.bstrVal, *end;
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
    
    for(i=0,end=buf+len; buf<end; i+=3)
    { for(j=0; j<4; j++)
      { do bin[j]=s_B64Dec[(U1)*buf++]; while(bin[j]!=255 && buf<end);
        if(buf==end) break;
      }
      if(bin[3]==64 || bin[2]==64)
      { if(bin[2] != 64)
        { ret[i  ]=(bin[0]<<2) | ((bin[1]>>4)&3);
          ret[i+1]=((bin[1]&0xF)<<4) | ((bin[2]>>2)&0xF);
        }
        else ret[i]=(bin[0]<<2) | ((bin[1]>>4)&3);
        break;
      }
      ret[i  ]=(bin[0]<<2) | ((bin[1]>>4)&3);
      ret[i+1]=((bin[1]&0xF)<<4) | ((bin[2]>>2)&0xF);
      ret[i+2]=((bin[2]&3)<<6) | bin[3];
    }
    
    out.Clear();
    out.v.vt=VT_ARRAY|VT_UI1, out.v.parray=sa;
  }
  else if(!b8bit && vIn.vt != VT_BSTR) return E_INVALIDARG;
  else
  { U1 *buf, *end;
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

    for(i=0,end=buf+len; buf<end; i+=3)
    { for(j=0; j<4; j++)
      { do bin[j]=s_B64Dec[*buf++]; while(bin[j]!=255 && buf<end);
        if(buf==end) break;
      }
      if(bin[3]==64 || bin[2]==64)
      { if(bin[2] != 64)
        { ret[i  ]=(bin[0]<<2) | ((bin[1]>>4)&3);
          ret[i+1]=((bin[1]&0xF)<<4) | ((bin[2]>>2)&0xF);
        }
        else ret[i]=(bin[0]<<2) | ((bin[1]>>4)&3);
        break;
      }
      ret[i  ]=(bin[0]<<2) | ((bin[1]>>4)&3);
      ret[i+1]=((bin[1]&0xF)<<4) | ((bin[2]>>2)&0xF);
      ret[i+2]=((bin[2]&3)<<6) | bin[3];
    }
    
    out.Clear();
    out.v.vt=VT_ARRAY|VT_UI1, out.v.parray=sa;
  }
  return S_OK;
} /* g_DecodeB64 */

/*** message hashing ***/
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
