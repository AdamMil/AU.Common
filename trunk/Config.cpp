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

// Config.cpp : Implementation of CConfig

#include "stdafx.h"
#include "Config.h"

// CConfig

STDMETHODIMP CConfig::get_Item(BSTR sKey, BSTR sType, BSTR sSection, VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;
  AXMLNode node = FindNode(sKey, sSection);
  if(node == NULL)
  { pvOut->vt=VT_NULL;
    return S_FALSE;
  }

  ASTR     type = sType&&sType[0] ? sType : node.Attr(ASTR(L"type"));
  AVAR     data;
  HRESULT  hRet = S_OK;

  data.Attach(node.Value().Detach());
  if(!type.IsEmpty())
  { if(sType==L"string")      hRet=data.ChangeType(VT_BSTR);
    else if(sType==L"int")    hRet=data.ChangeType(VT_I4);
    else if(sType==L"date")   hRet=data.ChangeType(VT_DATE);
    else if(sType==L"bigint") hRet=data.ChangeType(VT_I8);
    else if(sType==L"base64") ; // TODO: add base64 decoder
    else if(sType==L"hex")    ; // TODO: add hex decoder
  }
  if(SUCCEEDED(hRet)) *pvOut = data.Detach();
  return hRet;
} /* get_Item */

STDMETHODIMP CConfig::get_Exists(BSTR sKey, BSTR sSection, VARIANT_BOOL *pvBool)
{ if(!pvBool) return E_POINTER;
  *pvBool = FindNode(sKey, sSection) != NULL ? VBTRUE : VBFALSE;
  return S_OK;
} /* get_Exists */

STDMETHODIMP CConfig::get_Section(BSTR *psSect)
{ if(!m_XML)  return E_UNEXPECTED;
  if(!psSect) return E_POINTER;
  *psSect = m_Section.ToBSTR();
  return S_OK;
} /* get_Section */

STDMETHODIMP CConfig::put_Section(BSTR sSect)
{ if(!m_XML) return E_UNEXPECTED;
  AXMLNode node = m_XML.SelectSingleNode(sSect);
  if(node == NULL) return E_INVALIDARG;
  m_SectXML=node, m_Section=sSect;
  return S_OK;
} /* put_Section */

STDMETHODIMP CConfig::OpenXML(BSTR sXML)
{ Close();
  AComPtr<IXMLDOMDocument2> doc;
  HRESULT      hRet;
  VARIANT_BOOL succ;

  CHKRET(CREATE(DOMDocument, IXMLDOMDocument2, doc));
  CHKRET(doc->loadXML(sXML, &succ));
  if(!succ) return E_FAIL;
  m_XML = doc;
  return S_OK;
} /* OpenXML */

STDMETHODIMP CConfig::OpenFile(BSTR sPath)
{ Close();
  AComPtr<IXMLDOMDocument2> doc;
  HRESULT      hRet;
  VARIANT      var;
  VARIANT_BOOL succ;

  var.vt=VT_BSTR, var.bstrVal=sPath;
  CHKRET(CREATE(DOMDocument, IXMLDOMDocument2, doc));
  CHKRET(doc->load(var, &succ));
  if(!succ) return E_FAIL;
  m_XML = doc;
  return S_OK;
} /* OpenFile */

STDMETHODIMP CConfig::Close()
{ m_SectXML.Release();
  m_XML.Release();
  m_Section.Clear();
  return S_OK;
} /* Close */

AXMLNode CConfig::FindNode(BSTR sKey, BSTR sSection)
{ AXMLNode node, sect = sSection && sSection[0] ? m_XML.SelectSingleNode(sSection) : m_SectXML;
  if(sect == NULL || ((node=sect.SelectSingleNode(sKey))==NULL) && sect.NodeName() != L"Default")
    sect = m_XML.SelectSingleNode(ASTR(L"Default")), node = sect.SelectSingleNode(sKey);
  return node;
} /* FindNode */