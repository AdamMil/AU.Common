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

// Upload.cpp : Implementation of CUpload

#include "stdafx.h"
#include "Upload.h"
#include "UpFile.h"

#include <malloc.h>

// CUpload

/* ~(MODULES::UPLOAD, p'Upload::Form
  <PRE>[propget] HRESULT Form([in] BSTR sKey, [out,retval] VARIANT *pvOut);</PRE>
  The Form property retrieves the value from the form data. If the key is
  found, the return type will be a string (VT_BSTR). If the key is not found, the
  return type will be VT_NULL (null or Nothing). The key name match is case insensitive.
)~ */
STDMETHODIMP CUpload::get_Form(BSTR sKey, VARIANT *pvOut)
{ if(!pvOut) return E_POINTER;

  ASTR index(sKey);
  index.LowerCase();

  Map::const_iterator i = m_mForm.find(index);
  if(i == m_mForm.end())
  { pvOut->vt = VT_NULL;
    return S_FALSE;
  }

  const Item &item = (*i).second;  
  pvOut->vt = VT_BSTR;
  pvOut->bstrVal = ASTR::FromCStr((char*)item.Data, item.Len).Detach();
  return S_OK;
} /* get_Form */

/* ~(MODULES::UPLOAD, p'Upload::File
  <PRE>[propget] HRESULT File([in] VARIANT vIndex, [out,retval] IUpFile **ppFile);</PRE>
  The File property provides a method to access the files that were uploaded.
  If the type of 'vIndex' is VT_BSTR, it finds the file with the given form field name
  (case insensitive match). If the search fails, NULL (null or Nothing) is returned.
  Otherwise, if the type can be converted to VT_I4, it is used as an index into the
  files uploaded. If the index is out of range, an error will occur. The return type
  of this property is an `UpFile' object representing the file.
)~ */
STDMETHODIMP CUpload::get_File(VARIANT vIndex, IUpFile **ppFile)
{ if(!ppFile) return E_POINTER;

  Map::const_iterator i;

  if(vIndex.vt == VT_BSTR)
  { ASTR item(vIndex.bstrVal);
    item.LowerCase();
    i = m_mFiles.find(item);
    if(i == m_mFiles.end())
    { ppFile = NULL;
      return S_FALSE;
    }
  }
  else
  { AVAR var(vIndex);
    HRESULT hRet;
    CHKRET(var.ChangeType(VT_I4));
    if(var.v.intVal<0 || var.v.intVal>=(long)m_mFiles.size()) return E_INVALIDARG;
    i = m_mFiles.begin();
    while(var.v.intVal--) i++;
  }

  AComPtr<CUpFile> file;
  HRESULT hRet;
  CHKRET(CREATE(UpFile, IUpFile, file));
  file->Load(this, &(*i).second);
  file.CopyTo(ppFile);
  return S_OK;
} /* get_File */

/* ~(MODULES::UPLOAD, p'Upload::NumFiles
  <PRE>[propget] HRESULT NumFiles([out,retval] long *pnFiles);</PRE>
  The NumFiles property returns the number of files uploaded. This can be used in conjunction
  with the `File' property to iterate through the uploaded files.
)~ */
STDMETHODIMP CUpload::get_NumFiles(long *pnFiles)
{ if(!pnFiles) return E_POINTER;
  *pnFiles = (long)m_mFiles.size();
  return S_OK;
} /* get_NumFiles */

/* ~(MODULES::UPLOAD, p'Upload::NumFields
  <PRE>[propget] HRESULT NumFields([out,retval] long *pnFields);</PRE>
  The NumFields property returns the number of form fields. This can be used in conjunction
  with the `FieldKey' propety to iterate through the form fields.
)~ */
STDMETHODIMP CUpload::get_NumFields(long *pnFields)
{ if(!pnFields) return E_POINTER;
  *pnFields = (long)m_mForm.size();
  return S_OK;
} /* get_NumFields */

/* ~(MODULES::UPLOAD, p'Upload::FieldKey
  <PRE>[propget] HRESULT FieldKey([in] long nIndex, [out,retval] BSTR *psField);</PRE>
  The FieldKey property returns the name of a form field, given its index. If the index
  is out of range, an error will occur. The number of fields can be found using the
  `NumFields' property.
)~ */
STDMETHODIMP CUpload::get_FieldKey(long nIndex, BSTR *psField)
{ if(nIndex<0 || nIndex>=(long)m_mForm.size()) return E_INVALIDARG;
  if(!psField) return E_POINTER;

  Map::const_iterator i = m_mForm.begin();
  while(nIndex--) i++;
  *psField = (*i).first.ToBSTR();
  return S_OK;
} /* get_FieldKey */

#ifdef NDEBUG
  #pragma optimize("s", off)
  #pragma optimize("t", on)
#endif
STDMETHODIMP CUpload::OnStartPage(IDispatch *pDisp)
{ if(!pDisp) return E_POINTER;
  
  AComPtr<ASP::IScriptingContext> sc;
  AComPtr<ASP::IRequest>          req;
  VARIANT var;
  HRESULT hRet;

  CHKRET(DQUERYNS(pDisp, ASP, IScriptingContext, sc));
  CHKRET(sc->get_Request(&req));
  CHKRET(req->get_TotalBytes(&var.lVal));
  if(!var.lVal) return S_FALSE;

  var.vt=VT_I4;
  CHKRET(req->BinaryRead(&var, &m_vData));
  
  if(m_vData.Type() & VT_ARRAY)
  { SAFEARRAY *arr = m_vData.v.parray;
    // make sure array exists and is large enough
    if(var.lVal < 2 || !arr || arr->cbElements*arr->rgsabound[0].cElements < (ULONG)var.lVal) goto Done;
    char *delim, *data=(char*)arr->pvData, *end=data+var.lVal, *p, *q, *an, *ae;
    UA4   dlen;
    Item  item;
    bool  bFile;

    // first, seek to beginning of delimiter (should be there already...)
    while(data<end && isspace(*data)) data++;

    // now, determine post type
    if(memcmp(data, "--", 2)) // if not multipart form data
    { // assume it's www-form-urlencoded
      while(data<end)
      { p=(char*)memchr(data, '&', end-data);
        if(!p) p=end;
        q=(char*)memchr(data, '=', end-data);
        if(!q) break;
        *q++ = 0;
        item.Name.SetCStr(data);
        item.Data = (U1*)q;
        item.Len  = URLDecode(q, p);
        m_mForm.insert(MapPair(item.Name, item));
        data = p+1;
      }
    }
    else
    { // grab delimiter into delim
      p = (char*)memchr(data, '\r', end-data);
      if(!p) goto Done;

      // delimiter is preceded by a blank line
      p+=2, dlen=(UA4)(p-data);
      delim = (char*)alloca(dlen+1);
      delim[0]='\r', delim[1]='\n';
      memcpy(delim+2, data, dlen-2);
      delim[dlen]=0;

      data=p; // 'data' points at headers
      while(data<end)
      { bFile = false; // assume it's not a file
        item.ContentType.Clear(); // clear content type

        // process headers, break on a blank line
        while(data<end && *data != '\r')
        { // 'data' points at a header, find header name
          p = (char*)memchr(data, ':', end-data);
          if(!p) goto Done;
          *p++=0; // null terminate header name
          strlwr(data); // and convert it to lowercase
          while(p<end && isspace(*p)) p++; // move 'p' to beginning of header data
          q = (char*)memchr(p, '\r', end-p); // 'q' is the end of the line
          if(!q) goto Done;
          *q = 0;
          
          if(!strcmp(data, "content-disposition"))
          { data = p;
            while(data<q)
            { ae=(char*)memchr(data, ';', q-data); // find attribute end
              if(ae) *ae=0; else ae=q;
              p=strchr(data, '=');
              if(p) // if it doesn't have an equals sign, we're not interested
              { *p=0; strlwr(an=data); // null terminate and lowercase name 'an'
                data = ++p; // data points at attribute data
                if(*data=='\"') // quoted data, NOTE: this does not handle embedded quotes!
                { data=++p;
                  while(p<q && *p != '\"') p++;
                }
                else while(p<q && !isspace(*p)) p++; // read til space
                *p = 0; // null terminate data
                if(!strcmp(an, "name")) item.Name.SetCStr(data); // item name
                else if(!strcmp(an, "filename"))
                { item.FileName.SetCStr(data); // file name
                  bFile = true;
                }
              }
              data = ae+1; // move past end of attribute
              while(data<q && isspace(*data)) data++; // and seek to next attribute
            }
          }
          else if(!strcmp(data, "content-type")) item.ContentType.SetCStr(p);
          data = q+2; // move to next line
        }
        item.Data = (U1*)(data+=2); // save start of data
        if(data>=end) break;

        // attempt to find delimiter
        do
        { data=(char*)memchr(data, delim[0], end-data);
          if(!data || end-data < (IA4)dlen) goto Done;
          if(!memcmp(data, delim, dlen))
          { p = data + dlen;
            if(p[0]=='-'  && p[1]=='-')  p+=2;  // if it's the last one, advance two more
            if(p[0]=='\r' && p[1]=='\n') break; // make sure this is the end of a line
            data += dlen;
          }
          else data++;
        } while(data<end);
        if(data>=end) break;
        // p points to next section

        item.Len = (UA4)((U1*)data-item.Data); // calculate length from difference
        item.Name.LowerCase(); // lowercase the item name for case insensitivity
        if(bFile)
        { if(item.ContentType.IsEmpty()) item.ContentType=L"application/octet-stream";
          m_mFiles.insert(MapPair(item.Name, item));
        }
        else m_mForm.insert(MapPair(item.Name, item));
        
        data = p+2; // data points to line after delimiter
      }
    }
  }
  Done:
  return S_OK;
} /* OnStartPage */

UA4 CUpload::URLDecode(char *start, char *end)
{ char *write = start;
  UA4  len=(UA4)(end-start);
  char c, t;
  while(start<end)
  { c = *start;
    if(c=='+') c=' ';
    else if(c=='%')
    { if(end-start>=2)
      { t = tolower(*++start);
        if(t<'0' || t>'f' || (t>'9' && t<'a')) break;
        c = (t<'a' ? t-'0' : t-'a'+10) << 4;
        t = tolower(*++start);
        if(t<'0' || t>'f' || (t>'9' && t<'a')) break;
        c |= (t<'a' ? t-'0' : t-'a'+10);
        len-=2;
      }
    }
    *write++=c, start++;
  }
  return len;
}
#ifdef NDEBUG
  #pragma optimize("", on)
#endif
