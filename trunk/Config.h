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

// Config.h : Declaration of the CConfig

#pragma once
#include "resource.h"       // main symbols

#include "Common.h"


// CConfig

/* ~(MODULES::CONFIG, c'Config
  The Config object (ProgID AU.Common.Config) provides a method to easily access
  configuration information. The information is stored as XML (see `Schema')
  with the ability to embed type information. The Config object organizes data
  into Sections. Each section contains Keys, which hold the actual data. If a
  section doesn't have a particular key, the Config object will fall back to
  trying a section called "Default", where the default key values should be
  placed. First, Open*() should be called to open the configuration. Then,
  the other methods can be used. The Config object uses the "both" threading
  model.
)~ */

/* ~(MODULES::CONFIG, t'Config::Schema
  The XML schema is like this (lowercase tag names are constants, uppercase parts
  are to be defined by you, bracketed parts are optional):
  <PRE>
  <config>
    <SECTION_NAME>
      <FOO>
        <KEY [type="string|int|date|bigint|base64|hex"]>DATA</KEY>
      </FOO>
      <BAR>
        <KEY [type="string|int|date|bigint|base64|hex"]>DATA</KEY>
      </BAR>
      <KEY [type="string|int|date|bigint|base64|hex"]>DATA</KEY>
    </SECTION_NAME>
    <SECTION_NAME>...</SECTION_NAME>
    ...
  </config>
  
  An explanation of the types:
  string - the value is considered to be a string. The Text property of the XML
           node is used to retrieve the value. This is the default type, if the
           type is not specified.
  int    - the value is a 32-bit signed integer
  bool   - the value is a boolean, either 0/1 or true/false
  date   - the value is a date, with an optional time. The format can be
           mm/dd/yyyy hh:mm:ss, or any value that can be converted to a date
           by the system routines.
  bigint - the value is a 64-bit signed integer
  base64 - the value is a base64 encoded block of binary data
  hex    - the value is a hex encoded block of binary data
  If no type is specified, the Value property of the XML node is used to
  retrieve the value.
  </PRE>
  There is a special section called "Default", which you may include. If a key
  cannot be found in a requested section, the Config object will attempt to find
  it in the Default section. This allows you to specify default values for keys
  that aren't explicitly specified in other sections. As you can see from FOO
  and BAR, keys don't have to be immediate children of their section. They can
  be nested to any depth. This allows better organization of keys. Remember that
  XML is case sensitive, so the Config object is, too. Here is an example
  config file:
  <PRE>
  <config>
    <Default>
      <DB>
        <Default>CONNSTR1</Default>
        <fe_server>CONNSTR2</fe_server>
        <be_server>CONNSTR3</be_server>
      </DB>
      <website>www.mysite.com</website>
      <doom_day type="date">5/5/2005</doom_day>
      <random_bugs type="bool">0</random_bugs>
    </Default>

    <WebSite1>
      <DB>
        <be_server>CONNSTR4</be_server>
      </DB>
      <website>www.yoursite.com</website>
      <random_bugs type="bool">True</random_bugs>
    </WebSite1>
  </config>
  </PRE>
  In the preceding config file, a request for the key "DB/be_server" from section
  "WebSite1" would get "CONNSTR4", while a request for "DB/fe_server" from the same
  section (WebSite1) would get "CONNSTR2", since it fell back to the Default section
  upon not finding a "DB/fe_server" under WebSite1.
)~ */

class ATL_NO_VTABLE CConfig : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CConfig, &CLSID_Config>,
	public IDispatchImpl<IConfig, &IID_IConfig, &LIBID_AU_CommonLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
  DECLARE_REGISTRY_RESOURCEID(IDR_CONFIG)
  DECLARE_NOT_AGGREGATABLE(CConfig)
  BEGIN_COM_MAP(CConfig)
    COM_INTERFACE_ENTRY(IConfig)
    COM_INTERFACE_ENTRY(IDispatch)
    COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
  END_COM_MAP()
  DECLARE_PROTECT_FINAL_CONSTRUCT()
  DECLARE_GET_CONTROLLING_UNKNOWN()

	HRESULT FinalConstruct()
	{ return CoCreateFreeThreadedMarshaler(GetControllingUnknown(), &m_pUnkMarshaler.p);
	}

	void FinalRelease()
	{ m_pUnkMarshaler.Release();
	}

	CComPtr<IUnknown> m_pUnkMarshaler;

public:
  // IConfig
  STDMETHODIMP get_Item(/*[in]*/ BSTR sKey, /*[in,defaultvalue("")]*/ BSTR sType, /*[in,defaultvalue("")]*/ BSTR sSection, /*[out,retval]*/ VARIANT *pvOut);
  STDMETHODIMP get_Exists(/*[in]*/ BSTR sKey, /*[in,defaultvalue("")]*/ BSTR sSection, /*[out,retval]*/ VARIANT_BOOL *pbExists);
  STDMETHODIMP get_Section(/*[out,retval]*/ BSTR *psSect);
  STDMETHODIMP put_Section(/*[in]*/ BSTR sSect);

  STDMETHODIMP OpenXML(/*[in]*/ BSTR sXML);
  STDMETHODIMP OpenFile(/*[in]*/ BSTR sPath);
  STDMETHODIMP Close();
  
protected:
  AXMLNode m_XML, m_SectXML;
  ASTR     m_Section;
  
  AXMLNode FindNode(BSTR sKey, BSTR sSection);
};

OBJECT_ENTRY_AUTO(__uuidof(Config), CConfig)
