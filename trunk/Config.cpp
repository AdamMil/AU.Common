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

// Config.cpp : Implementation of CConfig

#include "stdafx.h"
#include "Config.h"

// CConfig

/* ~(MODULES::CONFIG, p'Config::Item
  <PRE>[defprop,propget] HRESULT Item([in] BSTR sKey,
                               [in,defaultvalue("")] BSTR sType,
                               [in,defaultvalue("")] sSection,
                               [out,retval] VARIANT *pvOut);</PRE>
  
  The Item property retrieves the value from the key given in 'sKey'.
  If 'sType' is passed, the return value will be forced to the given
  type if possible. If the type cannot be converted, the call will
  fail. If 'sType' is an empty string, the type specified on the key
  will be used. The types are described in `Schema'. If 'sSection' is
  passed, Item will attempt to read from that section. Otherwise, it
  will attempt to read from the current section, as set by `Section'.
  If the read attempt fails, it will fall back to the Default section,
  if there is one. If the key does not exist and sType is not specified,
  the function will not fail, but will instead return NULL (a variant
  with a type of VT_NULL).
  Item is also the default property, so in many languages, it can be
  used like this:
  <PRE>oConfig(parms) instead of oConfig.Item(parms)</PRE>
)~ */
STDMETHODIMP CConfig::get_Item(BSTR sKey, BSTR sType, BSTR sSection, VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  AComLock lock(this);
  AXMLNode node = FindNode(sKey, sSection);
  if(node == NULL)
  { if(sType&&sType[0]) return E_FAIL;
    pvOut->vt=VT_NULL;
    return S_FALSE;
  }

  ASTR     type = sType&&sType[0] ? sType : node.Attr(ASTR(L"type"));
  AVAR     data;
  HRESULT  hRet = S_OK;

  data.Attach(node.Value().Detach());
  if(!type.IsEmpty())
  { if(sType==L"string")      hRet=data.ChangeType(VT_BSTR);
    else if(sType==L"int")    hRet=data.ChangeType(VT_I4);
    else if(sType==L"bool")   hRet=data.ChangeType(VT_BOOL);
    else if(sType==L"date")   hRet=data.ChangeType(VT_DATE);
    else if(sType==L"bigint") hRet=data.ChangeType(VT_I8);
    else if(sType==L"base64") hRet=g_DecodeB64(data, true, data);
    else if(sType==L"hex")    hRet=g_HexToBin(data.ToBSTR(), data);
  }
  if(SUCCEEDED(hRet)) *pvOut = data.Detach();
  return hRet;
} /* get_Item */

/* ~(MODULES::CONFIG, p'Config::Exists
  <PRE>[propget] HRESULT Exists([in] BSTR sKey, [in,defaultvalue("")] sSection,
                         [out,retval] VARIANT_BOOL *pbExists);</PRE>
  The Exists property attempts to find the specified key (using the same semantics
  as `Item'). It returns true (nonzero) if the key is found (in either the
  specified/current section OR the Default section) and false (zero) if it's not
  found.
)~ */
STDMETHODIMP CConfig::get_Exists(BSTR sKey, BSTR sSection, VARIANT_BOOL *pbExists)
{ if(!pbExists) return E_POINTER;
  AComLock lock(this);
  *pbExists = FindNode(sKey, sSection) != NULL ? VBTRUE : VBFALSE;
  return S_OK;
} /* get_Exists */

/* ~(MODULES::CONFIG, p'Config::Section
  <PRE>
  [propget] HRESULT Section([out,retval] BSTR *psSect);
  [propput] HRESULT Section([in] BSTR sSect);
  </PRE>
  The Section property sets or retrieves the current section. The Config object will
  attempt to look in the current section before looking in the Default section. The
  default value for the current section is an empty string, which means no section
  is the current section.
)~ */
STDMETHODIMP CConfig::get_Section(BSTR *psSect)
{ if(!psSect) return E_POINTER;
  AComLock lock(this);
  if(!m_XML)  return E_UNEXPECTED;
  *psSect = m_Section.ToBSTR();
  return S_OK;
} /* get_Section */

STDMETHODIMP CConfig::put_Section(BSTR sSect)
{ AComLock lock(this);
  if(!m_XML) return E_UNEXPECTED;
  AXMLNode node = m_XML.SelectSingleNode(sSect);
  if(node == NULL) return E_INVALIDARG;
  m_SectXML=node, m_Section=sSect;
  return S_OK;
} /* put_Section */
  
/* ~(MODULES::CONFIG, f'Config::OpenXML
  <PRE>HRESULT OpenXML([in] BSTR sXML);</PRE>
  The OpenXML method loads the Config object using an XML string. The XML string
  is passed in the 'sXML' argument. `Close' will be called automatically, regardless
  of whether or not the configuration could be successfully loaded.
)~ */
STDMETHODIMP CConfig::OpenXML(BSTR sXML)
{ AComLock lock(this);
  Close();
  AComPtr<IXMLDOMDocument2> doc;
  HRESULT      hRet;
  VARIANT_BOOL succ;

  CHKRET(CREATE(DOMDocument, IXMLDOMDocument2, doc));
  CHKRET(doc->loadXML(sXML, &succ));
  if(!succ) return E_FAIL;
  return doc->get_firstChild(&m_XML);
} /* OpenXML */

/* ~(MODULES::CONFIG, f'Config::OpenFile
  <PRE>HRESULT OpenFile([in] BSTR sPath);</PRE>
  The OpenFile method loads the Config object from a file. The path to the file
  is passed in the 'sPath' argument. The path should be absolute, but can be local
  (ie. "C:\file.xml"), on a share (ie. "\\server\share\file.xml"), or on a web site
  (ie. "http://www.boo.com/config.xml"). It is recommended that the file be local
  or on a share, though, both for performance reasons and to prevent the load from
  failing due to the server being busy, etc. `Close' will be called automatically,
  regardless of whether or not the configuration could be successfully loaded.
)~ */
STDMETHODIMP CConfig::OpenFile(BSTR sPath)
{ AComLock lock(this);
  Close();
  AComPtr<IXMLDOMDocument2> doc;
  HRESULT      hRet;
  VARIANT      var;
  VARIANT_BOOL succ;

  var.vt=VT_BSTR, var.bstrVal=sPath;
  CHKRET(CREATE(DOMDocument, IXMLDOMDocument2, doc));
  CHKRET(doc->load(var, &succ));
  if(!succ) return E_FAIL;
  return doc->get_firstChild(&m_XML);
} /* OpenFile */

/* ~(MODULES::CONFIG, f'Config::Close
  <PRE>HRESULT Close();</PRE>
  The Close method closes the configuration. Calling this is not necessary, as the
  configuration will be closed when it is no longer needed.
)~ */
STDMETHODIMP CConfig::Close()
{ AComLock lock(this);
  m_SectXML.Release();
  m_XML.Release();
  m_Section.Clear();
  return S_OK;
} /* Close */

AXMLNode CConfig::FindNode(BSTR sKey, BSTR sSection)
{ AXMLNode node, sect = sSection && sSection[0] ? m_XML.SelectSingleNode(sSection) : m_SectXML;
  if(sect == NULL || ((node=sect.SelectSingleNode(sKey))==NULL) && sect.NodeName() != L"Default")
  { sect = m_XML.SelectSingleNode(ASTR(L"Default"));
    if(sect != NULL) node = sect.SelectSingleNode(sKey);
  }
  return node;
} /* FindNode */
