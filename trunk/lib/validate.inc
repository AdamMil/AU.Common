<script runat="server" language="jscript">
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

var _VAL_DESC,  _VAL_MAXLEN, _VAL_MINLEN, _VAL_TYPE,  _VAL_CLEAN,  _VAL_NOMATCH,
    _VAL_NMMSG, _VAL_MATCH,  _VAL_MMSG,   _VAL_RANGE, _VAL_EXECAV, _VAL_EXECBV;
var _VAL_AttrNames, _VAL_bInit, _VAL_bWroteCS;

function _VAL_Init()
{ _VAL_DESC =0, _VAL_MAXLEN=1, _VAL_MINLEN=2, _VAL_TYPE =3, _VAL_CLEAN =4,  _VAL_NOMATCH=5;
  _VAL_NMMSG=6, _VAL_MATCH =7, _VAL_MMSG  =8, _VAL_RANGE=9, _VAL_EXECAV=10, _VAL_EXECBV =11;

  _VAL_AttrNames = new Array("desc", "maxlength", "minlen", null, "clean", "nomatch", "nmmsg", "match", "mmsg", "range", "execav", "execbv");
  _VAL_bInit     = true;

  if(!UA_bLoaded) UA_Init();
}

function Validator(form_id)
{ if(typeof(UA_bLoaded)=="undefined") throw "Error: UA_Init() from browser.inc must be called before this method";
  if(!_VAL_bInit) _VAL_Init();

  this.Add      = _VAL_Add;
  this.Control  = _VAL_Control;
  this.Validate = _VAL_Validate;
  this.Write    = _VAL_Write;
  this.Field    = _VAL_Field;
  this.OnSubmitCode = _VAL_OnSubmitCode;
  this.WriteCSCode  = _VAL_WriteCSCode;

  this.FormID   = form_id;
  this.Controls = new Array();
  this.CtlIndex = new Array();
  this.Message  = null;
}

function _VAL_ControlObj(par, name, desc, maxlen, minlen, type, value, clean, nomatch, nmmsg, match, mmsg, range, execav, execbv)
{ this.Validate = _VALCTL_Validate;
  this.HTML     = _VALCTL_HTML;
  this.Field    = _VALCTL_Field;

  this.Parent = par;
  this.Name   = name;
  this.Data   = new Array(desc, maxlen, minlen, type, clean, nomatch, nmmsg, match, mmsg, range, execav, execbv);
  this.Value  = value ? value : Request.Form(name)>"" ? this.Field(name) : null;
}

function _VAL_Add(name, desc, maxlen, minlen, type, clean, value, nomatch, nmmsg, match, mmsg, range, execav, execbv)
{ if(!name) throw "Field must have a name!";
  name = String(name).toLowerCase();
  if(type) type = String(type).toLowerCase();
  var ctl = new _VAL_ControlObj(this, name, desc, maxlen, minlen, type, value, clean, nomatch, nmmsg, match, mmsg, range, execav, execbv);
  this.Controls[this.Controls.length] = ctl;
  this.CtlIndex[name] = ctl;
  return ctl;
}

function _VAL_Control(name)
{ var ctl = this.CtlIndex[name];
  return ctl ? ctl : null;
}

function _VAL_Validate()
{ var i=0, arr=this.Controls, len=arr.length;
  for(i=0; i<len; i++) if(!arr[i].Validate()) return false;
  return true;
}

function _VAL_Write(name, attr, value)
{ var ctl = this.CtlIndex[name];
  return ctl ? ctl.HTML(attr, value) : "ERROR: Unregistered control "+name;
}

function _VAL_Field(name)
{ ctl = this.Control(name);
  return ctl ? ctl.Value : "";
}

function _VAL_OnSubmitCode()
{ return UA_Level>=1 ? "return VAL_Validate(this);" : "";
}

function _VAL_WriteCSCode()
{ if(!this.Controls.length) return;
  if(!_VAL_bWroteCS)
  { Response.Write("\n<\script src=\"/js/validate.js\"><\/script>");
    if(UA_Level>=2) Response.Write("<\script src=\"/js/validateIE.js\"><\/script>\n");
    else Response.Write("<\script src=\"/js/validateNS.js\"><\/script>\n");
  }

  var i, ctls=this.Controls, len=ctls.length, bNeedInit=false;

  if(UA_Level==1)
  { Response.Write("\n<\script>\n");
    Response.Write("if(typeof(_VAL_Forms)==\"undefined\") { var _VAL_Forms=new Array(); }\n");
    Response.Write("_VAL_Forms[\""); Response.Write(this.FormID); Response.Write("\"]=new Array(\n");
    var j, ctl, type, dat, dlen;
    for(i=0; i<len; i++)
    { ctl=ctls[i], dat=ctl.Data, dlen=dat.length;
      if(i) Response.Write(",\n");
      Response.Write("  new Array(\""); Response.Write(ctl.Name); Response.Write('\"')
      for(j=0; j<dlen; j++)
      { Response.Write(", ");
        if(dat[j]==null) Response.Write("null");
        else if(typeof(dat[j])=="number") Response.Write(dat[j]);
        else
        { Response.Write('\"'); Response.Write(_VAL_Escape(dat[j])); Response.Write('\"');
        }
      }
      Response.Write(')');
      
      if(!bNeedInit && dat[_VAL_TYPE]=="select" && ctl.Value != null) bNeedInit=true;
    }
    Response.Write("\n);\n");
  }
  else if(UA_Level>=2)
  { for(i=0; i<len; i++)
      if(ctls[i].Data[_VAL_TYPE]=="select" && ctls[i].Value != null) { bNeedInit=true; break; }
  }

  if(bNeedInit)
  { if(UA_Level>=2) Response.Write("\n<\script>\n");
    Response.Write("HOOK_addOnLoad(_VAL_InitForm, ");
    Response.Write(this.FormID);
    Response.Write(");\n");
    Response.Write("<\/script>\n");
  }
  else if(UA_Level==1) Response.Write("<\/script>\n");

  _VAL_bWroteCS = true;
}

function _VAL_Escape(v)
{ v = String(v);
  v = v.replace(/\"/g, "\\\"");
  v = v.replace(/\n/g, "\\n");
  v = v.replace(/\t/g, "\\t");
  return v;
}

function _VALCTL_Validate()
{ var arr = this.Data, v, len, fun, regex, err=false, msg;

  this.Value = this.Value ? String(this.Value) : "";
  if((v=arr[_VAL_CLEAN]))
    this.Value = this.Value.replace(new RegExp(v, "gm"), "");
  if((v=arr[_VAL_EXECBV]))
  { fun = new Function("ctl", v);
    ret = fun(this);
    if(typeof(ret)=="string") err=true, msg=ret;
    else if(ret==false) return true;
  }
  if(!err && (v=arr[_VAL_MAXLEN]) && this.Value.length>v)
    err=true, msg="must be at most "+v+" characters.";
  if(!err && (v=arr[_VAL_MINLEN]))
  { len = this.Value.length;
    if(len < v)
      err=true, msg = (v==1 ? "is required." : "must be at least "+v+" characters.");
    else if(v<0 && len>0 && len<-v)
      err=true, msg = "must be at least "+(-v)+" characters, if specified.";
  }
  if(!err && (v=arr[_VAL_NOMATCH]))
  { if(this.Value.search(new RegExp(v, "gm")) != -1)
    { err=true, msg=this.Data[_VAL_NMMSG];
      if(!msg) msg="is invalid (matches "+v+").";
    }
  }
  if(!err && (v=arr[_VAL_MATCH]))
  { if(this.Value.search(new RegExp(v, "gm")) == -1)
    { err=true, msg=this.Data[_VAL_MMSG];
      if(!msg) msg="fails to match required format ("+v+").";
    }
  }
  if(!err && (v=arr[_VAL_RANGE]))
  { var val = parseFloat(this.Value);
    v = v.split(' ');
    if(val<parseFloat(v[0]) || val>parseFloat(v[1]))
      err=true, msg = "is outside of expected range ("+v[0]+" to "+v[1]+").";
  }
  if(err)
  { v = this.Data[_VAL_DESC];
    this.Parent.Message = (v ? v : this.Name)+' '+msg;
    return false;
  }
  else if((v=arr[_VAL_EXECAV]))
  { fun = new Function("ctl", v);
    fun(this);
  }
  return true;
}

function _VALCTL_HTML(attr, value)
{ var type = this.Data[_VAL_TYPE], ret;

  if(type==null) type="text";
  if(type=="textbox") ret="<textbox";
  else if(type=="select") ret="select";
  else
  { ret="<input";
    if(type!="text") ret += " type=\"", ret += type, ret += '\"';
  }
  ret += " name=\"", ret += this.Name, ret += '\"';

  if(UA_Level>=2)
  { var fn;
    for(i=0; i<_VAL_AttrNames.length; i++)
    { if((fn=_VAL_AttrNames[i]) && this.Data[i])
        ret += ' ', ret += fn, ret += "=\"", ret += this.Data[i], ret += '\"';
    }
  }

  if(this.Data[_VAL_MAXLEN] && type=="textarea" && UA_Level>=1)
  { ret += " onchange=\"_VAL_CheckLen(this";
    if(UA_Level==1) ret += ", ", ret += this.Data[_VAL_MAXLEN];
    ret += ")\"";
  }
  
  if(attr) { ret += ' '; ret+=attr; }

  if(type != "password")
  { if(value) this.Value = value;
    if(this.Value)
    { if(type=="textbox")
      { ret += '>'; ret += this.Value; ret += "</textbox>";
        return ret;
      }
      else if(type != "select")
      { ret += " value=\""; ret += this.Value; ret += "\"/>";
        return ret;
      }
    }
  }

  if(type=="textbox") return ret += "></textbox>";
  else if(type=="select") return ret += ">";
  else return ret += "/>";
}

function _VALCTL_Field()
{ var v=Request.Form(this.Name);
  if(v>"")
  { v = String(v);
    if(this.Data[_VAL_CLEAN])
    { var regex = new RegExp(this.Data[_VAL_CLEAN], "gm");
      v=v.replace(regex, "");
    }
    return v;
  }
  else return "";
}
</script>