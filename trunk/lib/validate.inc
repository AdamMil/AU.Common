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

/* ~(MODULES::INC::VALIDATE, t'validate.inc
  The validate.inc library provides validation support that will take advantage of
  client-side javascript support if it is available. It depends on the
  `BROWSER::browser.inc' library and the `JS::HOOKS::hooks.js' client libraries,
  as well as the validate.js, validateIE.js, and validateNS.js files. It assumes
  that the javsacript files are located in /js/. The validation library supports
  multiple forms and multiple validation types per form. The basic operation is
  that you create a Validator object for each validation set using `Validator'
  (usually one per form), and then you add controls to it using `Validator::Add'.
  You can call `Validator::Validate' to check the form at any time. To add the
  appropriate client-side code, you should use `Validator::WriteCSCode' and
  `Validator::OnSubmitCode'. The Validator object supports length checks,
  type checks, cleaning by regex, matching by regex, range checks, and checks by
  arbitrary javascript code.
)~ */
  
  
var _VAL_DESC,  _VAL_MAXLEN, _VAL_MINLEN, _VAL_TYPE,  _VAL_DTYPE,  _VAL_CLEAN,  _VAL_NOMATCH,
    _VAL_NMMSG, _VAL_MATCH,  _VAL_MMSG,   _VAL_RANGE, _VAL_EXECBV, _VAL_EXECAV;
var _VAL_AttrNames, _VAL_bInit, _VAL_bWroteCS;

function _VAL_Init()
{ _VAL_DESC =0, _VAL_MAXLEN=1, _VAL_MINLEN=2, _VAL_TYPE =3,  _VAL_DTYPE =4,  _VAL_CLEAN =5, _VAL_NOMATCH=6;
  _VAL_NMMSG=7, _VAL_MATCH =8, _VAL_MMSG  =9, _VAL_RANGE=10, _VAL_EXECBV=11, _VAL_EXECAV=12;

  _VAL_AttrNames = new Array("desc", "maxlength", "minlen", null, "dtype", "clean", "nomatch", "nmmsg",
                             "match", "mmsg", "range", "execbv", "execav");
  _VAL_bInit     = true;

  if(!UA_bLoaded) UA_Init();
}

/* ~(MODULES::INC::VALIDATE, c'Validator
  <PRE>function Validator(form_id);</PRE>
  The Validator constructor takes the contents of the 'id' attribute of the form and
  constructs a new Validator object. `BROWSER::browser.js' must be initialized prior
  to calling this, or an exception will be thrown. The Validator object has the
  following properties:
  <PRE>
  FormID   - a copy of the 'form_id' parameter passed to the constructor
  Controls - an array of ControlObj objects representing added controls, in
             the order they were added
  CtlIndex - an array of ControlObj objects indexed by lowercase field name
  Message  - the error message from an unsuccessful validation is stored here.
             if no error has occurred yet, this value is null. it is not a
             problem to write to this field.
  GetRawField   - a reference to a function used to read the raw form data.
  GetFieldCount - a reference to a function used to get the number of fields.</PRE>
)~ */
function Validator(form_id)
{ if(typeof(UA_bLoaded)=="undefined")
    _VAL_Die("UA_Init() from browser.inc must be called before this method");
  if(!_VAL_bInit) _VAL_Init();

  this.Add      = _VAL_Add;
  this.Control  = _VAL_Control;
  this.Validate = _VAL_Validate;
  this.Write    = _VAL_Write;
  this.Field    = _VAL_Field;
  this.RawField = _VAL_RawField;
  this.Persist  = _VAL_Persist;
  this.Restore  = _VAL_Restore;
  this.OnSubmitCode = _VAL_OnSubmitCode;
  this.WriteCSCode  = _VAL_WriteCSCode;
  
  this.GetRawField   = _VAL_GetRawField;
  this.GetFieldCount = _VAL_GetFieldCount;

  this.FormID   = form_id;
  this.Controls = new Array();
  this.CtlIndex = new Array();
  this.Message  = null;
}

/* ~(MODULES::INC::VALIDATE, c'ControlObj
  The ControlObj object represents a single control within a Validator. Generally,
  you will never need to use the ControlObj object directly, but it is provided for
  advanced purposes. The ControlObj object contains the following methods: 
  `Validate', `HTML', `Field', `SetVal'. In addition, it contains the following properties:
  <PRE>
  Parent  - a reference to the parent Validator object
  Name    - the form name of the control
  Data    - an array containing the validation parameters for this control, in the
           following order: desc, maxlen, minlen, type, dtype, clean, nomatch,
           nmmsg, match, mmsg, range, execav, execbv
  Value   - the value of the control. This should be set using SetVal. Unless manually
           changed, this will always be a string or null.
  Typed   - the cached typed value of the control. This is recalculated if null.
  Written - a boolean value that is true if the control has been written already. This
            does not necessarily indicate that the data has been sent to the client. It
            only indicates that HTML() has been called.</PRE>
)~ */
function _VAL_ControlObj(par, name, desc, maxlen, minlen, type, dtype, value, clean, nomatch, nmmsg,
                         match, mmsg, range, execbv, execav)
{ this.Validate = _VALCTL_Validate;
  this.HTML     = _VALCTL_HTML;
  this.Field    = _VALCTL_Field;
  this.SetVal   = _VALCTL_SetVal;

  this.Parent  = par;
  this.Name    = name;
  this.Data    = new Array(desc, maxlen, minlen, type, dtype, clean, nomatch, nmmsg, match, mmsg,
                           range, execbv, execav);
  if(type=="checkbox")
    this.Value = par.GetRawField(name)>"" ? "1" : par.GetFieldCount() ? null : value==null ? null : value;
  else
    this.Value = par.GetRawField(name)!=null ? this.Field(name) : value==null ? null : value; // squash 'undefined'
  this.Typed   = null;
  this.Written = false;
}

/* ~(MODULES::INC::VALIDATE, f'Validator::Add
  <PRE>ControlObj Add(name, desc, maxlen, minlen, type, dtype, value, clean,
       nomatch, nmmsg, match, mmsg, range, execbv, execav);</PRE>
  The Validator constructor takes the contents of the 'id' attribute of the form and
  constructs a new Validator object. `BROWSER::browser.js' must be initialized prior
  to calling this, or an exception will be thrown. The constructor takes a large number
  of parameters, which are described below. The only required parameter is the name.
  'null' can be passed for any other parameter to effectively bypass it. Regexp parameters
  are matched with the global and multi-line flags, and should contain only the regexp.
  For instance, /a[bc]d/ would be specified as "a[bc]d". Remember to escape backslashes
  in the regexp string as necessary!
  <PRE>
  name    - the name of the field in the form (the name attribute), ie "login"
  desc    - a short, user friendly description of the field, ie "Login Name"
  maxlen  - the maximum number of characters the raw field is allowed to be
  minlen  - the minimum number of characters the raw field is allowed to be.
            this can be set to 1 to make the field required, and if a
            negative number is passed, this sets the minimum length without
            making the field required. (ie, -5 will flag an error on lengths
            less than 5, unless the field is empty)
  type    - the type of the field (text, password, radio, checkbox, textarea,
            select, etc). if not passed or null, text is assumed
  dtype   - the data type of the field (string, int, float, date). the field
            must be able to be converted into the given type. if not passed
            or null, the default is "string", which essentially does no type
            conversion
  value   - the default value of the control. if the control cannot be
            initialized from form data, it will be initialized from 'value'
  clean   - this is a regexp which is used to clean the field data. any parts
            of the input data that match this regexp will be replaced with ""
            even the "raw" field data is cleaned. for instance "\D" will strip
            all non-digits from the input
  nomatch - the is a regexp that the field must not match. if the field matches
            this regexp, the validation fails
  nmmsg   - this is the message to be displayed if the 'nomatch' check fails.
            it should be in a form that assumes the field name will be
            prepended. for example, "cannot contain the word 'bill'.". if this
            is not passed, a default message will be used.
  match   - this is a regexp that the message must match.
  mmsg    - this is the message to be displayed if the 'match' check fails. see
            the 'nmmsg' parameter for more information
  range   - this is a pair of numbers separated by a single space. the first
            number should be less than or equal to the second number. the
            validation fails if the input is not a float that is within the
            range, inclusive. for instance, "-5 5" allows numbers from -5 to
            5. -5 and 5 would both be okay, but anything less than -5 or more
            than 5 would cause an error.
  execbv  - this is a snippet of javascript code that will be executed after
            cleaning takes place, but before any actual validation. the value
            is assumed to be the body of a function. the function takes the
            ControlObj as a parameter called 'ctl'. if the code returns a
            string, validation fails with the string as the error message,
            prepended by the field name, as always. if the code returns false,
            validation succeeds instantly. the subsequent checks are skipped,
            and execav is not called. any other return value (ie, true) will
            cause validation to continue as normal
  execav  - this is a snippet of javascript that will be executed after
            validation. it is expected to be the body of a function that takes
            three parameters, 'ctl', 'err', and 'msg'. 'ctl' is the current
            control. 'err' is true if validation has failed so far, and 'msg'
            is the error message if validation has failed so far. if the
            code returns a string, the validation will fail with the string
            as the error code. if the function returns false, validation
            will succeed, regardless of whether any checks had failed.
            any other return will have no effect.
  </PRE>
  Validation is performed in the following order, and if any step fails, the
  other steps are not executed (with the exception of 'execav', which always
  runs, unless 'execbv' returns false):
  clean, execbv, dtype, maxlen, minlen, nomatch, match, range, execav. The Add method
  returns the new `ControlObj' object.
)~ */
function _VAL_Add(name, desc, maxlen, minlen, type, dtype, value, clean, nomatch, nmmsg, match,
                  mmsg, range, execbv, execav)
{ if(!name) _VAL_Die("Field must have a name!");
  name = String(name).toLowerCase();
  if(type)  type  = String(type).toLowerCase();
  if(dtype) dtype = String(dtype).toLowerCase();
  var ctl = new _VAL_ControlObj(this, name, desc, maxlen, minlen, type, dtype, value, clean,
                                nomatch, nmmsg, match, mmsg, range, execbv, execav);
  this.Controls[this.Controls.length] = ctl;
  this.CtlIndex[name] = ctl;
  return ctl;
}

/* ~(MODULES::INC::VALIDATE, f'Validator::Control
  <PRE>ControlObj Control(name)</PRE>
  The Control method looks up a control by name and returns the `ControlObj' representing it.
  If the control cannot be found, null is returned.
)~ */
function _VAL_Control(name)
{ var ctl = this.CtlIndex[name.toLowerCase()];
  return ctl ? ctl : null;
}

/* ~(MODULES::INC::VALIDATE, f'Validator::Validate
  <PRE>bool Validate()</PRE>
  This method validates the controls in the order they were added. If any control fails
  validation, the method immediately returns false and does not validate the other
  controls. If all controls validate successfully, the method returns true.
)~ */
function _VAL_Validate()
{ var i=0, arr=this.Controls, len=arr.length;
  for(i=0; i<len; i++) if(!arr[i].Validate()) return false;
  return true;
}

/* ~(MODULES::INC::VALIDATE, f'Validator::Write
  <PRE>string Write(name[, attr][, value])</PRE>
  This method returns the HTML for the control specified by the given name.
  The HTML written will depend upon the browser type and the validation
  options set. For controls of type 'select', this writes the opening
  select tag, but does not write any options or the closing select tag.
  'attr' is an optional string containing additional HTML attriutes that
  will be inserted into the HTML tag. If 'value' is passed, it will
  be stored in the .Value property and be used in place of it when writing
  the value of the control. Passing null will not override the .Value
  property, as null indicates an unpassed option. To effectively nullify
  the .Value property, you can pass an empty string. An exception is made
  for radio controls. Since they generally are written multiple times,
  with a different value each time, this function will NOT store 'value'
  into the .Value property, but will instead check .Value against 'value'
  and add checked="1" to the HTML if the two match. Passing null for 'value'
  (or not passing it) will simply use .Value as the value for the radio control 
  If the control does not exist, a conspicuous-looking error message is
  returned as the HTML.
)~ */
function _VAL_Write(name, attr, value)
{ var ctl = this.Control(name);
  return ctl ? ctl.HTML(attr, value) : "ERROR: Unregistered control "+name;
}

/* ~(MODULES::INC::VALIDATE, f'Validator::Field
  <PRE>variant Field(name)</PRE>
  This method returns the properly-typed value for the control named. An exception will
  be thrown if the control does not exist or the control's value cannot be converted
  to the type specified in the 'dtype' parameter to `Add'. However, if the control's
  value is empty/null and a data type other than "string" was specified, it will not
  be converted, and null will be returned. Otherwise, if the control's value is
  empty/null an empty string will be returned.
  If the control is of type "checkbox", false will be returned if the field is empty,
  and true otherwise.
)~ */
function _VAL_Field(name)
{ ctl = this.Control(name);
  if(!ctl) _VAL_Die("No such control!");
  if(ctl.Typed != null) return ctl.Typed;

  var v = ctl.Value, type;
  if(ctl.Data[_VAL_TYPE]=="checkbox") v = v>"";
  else if((type=ctl.Data[_VAL_DTYPE]) != null && type != "string")
  { if(v==null || v=="") return null;
    v=_VAL_CoerceType(String(v), type);
    if(v==null) _VAL_Die(name+" could not be converted to type '"+type+"'. the value is: "+ctl.Value);
  }
  else if(v==null || v=="") v="";
  return ctl.Typed=v;
}

/* ~(MODULES::INC::VALIDATE, f'Validator::RawField
  <PRE>string RawField(name)</PRE>
  This method returns the mostly-raw value of the field. If the data came from the form, it
  will be cleaned by the 'clean' regex passed to `Add'. The return value is always a string.
  If the control's value is null, an empty string will be returned.
)~ */
function _VAL_RawField(name)
{ ctl = this.Control(name);
  if(!ctl) _VAL_Die("No such control!");
  return ctl.Value==null ? "" : ctl.Value;
}

/* ~(MODULES::INC::VALIDATE, f'Validator::Persist
  <PRE>string Persist()</PRE>
  This method saves the current values of all the controls into a string. The string can later
  be passed to `Restore' to restore the Validator to the state it was in when Persist() was
  called.
)~ */
function _VAL_Persist()
{ var i, ctl, ctls=this.Controls, len=ctls.length, ret="";
  for(i=0; i<len; i++)
  { ctl = ctls[i];
    if(i) ret+=',';
    ret += ctl.Name;
    ret += ',';
    if(ctl.Value==null) ret+="n";
    else
    { ret += "s";
      ret += String(ctl.Value).replace(/%/g, "%1").replace(/,/g, "%2");
    }
  }
  return ret;
}

/* ~(MODULES::INC::VALIDATE, f'Validator::Restore
  <PRE>void RawField(string)</PRE>
  This method takes a string returned from `Persist' and uses it to restore the state of the
  Validator to the state it was in when Persist() was called.
)~ */
function _VAL_Restore(dat)
{ dat = dat.split(',');
  var i, len=dat.length, val, c;
  for(i=0; i<len; i+=2)
  { ctl = this.Control(dat[i]);
    val = dat[i+1];
    c   = val.charAt(0);
    if(c=='n') ctl.Value=null;
    else if(c=='s') ctl.Value=val.substr(1).replace(/%2/g, ",").replace(/%1/g, "%");
    ctl.Typed = null;
  }
}

/* ~(MODULES::INC::VALIDATE, f'Validator::GetRawField
  <PRE>variant GetRawField(key)</PRE>
  This method returns the raw value of the field. Normally, it is only used internally by the
  validation code, but you can use it if you want. It returns either null or a string containing
  the value referenced by 'key'. If you have a custom method of field access (ie, from an upload
  control), you can create your own GetRawField function and assign it to the validator to force
  it to use your function instead. If you override this, you should probably override
  `GetFieldCount' as well.
)~ */
function _VAL_GetRawField(key)
{ var v = Request.Form(key);
  return v>"" ? v : null;
}

/* ~(MODULES::INC::VALIDATE, f'Validator::GetFieldCount
  <PRE>long GetFieldCount()</PRE>
  This method returns the number of fields actually submitted. If you override this, you should
  probably override `GetRawField' as well.
)~ */
function _VAL_GetFieldCount()
{ return Request.Form.Count;
}

/* ~(MODULES::INC::VALIDATE, f'Validator::OnSubmitCode
  <PRE>string OnSubmitCode(name)</PRE>
  This method returns a snippet of javascript that is designed to be the value of the
  form's onsubmit attribute. For browsers that support our client-side validation,
  some code will be written to perform that validation and abort the post if the
  validation failed. If this is used, `WriteCSCode' should be used as well.
)~ */
function _VAL_OnSubmitCode()
{ return UA_Level>=1 ? "VAL_SubmitBlur(this); return VAL_Validate(this);" : "";
}

/* ~(MODULES::INC::VALIDATE, f'Validator::WriteCSCode
  <PRE>void WriteCSCode(name)</PRE>
  This method writes the client-side code necessary to support this validator on the browser.
  It should be used in a location where &lt;script&gt; blocks could safely go. In addition,
  it should be called AFTER all form elements have been generated with `Write' or
  `ControlObj::HTML'. If it is not, 
)~ */
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
      
      if(!bNeedInit && ctl.Written && dat[_VAL_TYPE]=="select" && ctl.Value != null) bNeedInit=true;
    }
    Response.Write("\n);\n");
  }
  else if(UA_Level>=2)
  { for(i=0; i<len; i++)
      if(ctls[i].Data[_VAL_TYPE]=="select" && ctls[i].Written && ctls[i].Value != null) { bNeedInit=true; break; }
  }

  if(bNeedInit)
  { if(UA_Level>=2) Response.Write("\n<\script>\n");
    Response.Write("HOOK_addOnLoad(_VAL_InitForm, new Array(");
    for(i=0,j=false; i<len; i++)
    { ctl = ctls[i];
      if(ctl.Data[_VAL_TYPE]=="select" && ctl.Written && ctl.Value != null)
      { if(j) Response.Write(", ");
        else j=true;
        Response.Write(this.FormID);
        Response.Write('.');
        Response.Write(ctl.Name);
        Response.Write(",\"");
        Response.Write(_VAL_Escape(ctl.Value));
        Response.Write('\"');
      }
    }
    Response.Write("));\n");
    Response.Write("<\/script>\n");
  }
  else if(UA_Level==1) Response.Write("<\/script>\n");

  _VAL_bWroteCS = true;
}

function _VAL_Escape(v)
{ v = String(v);
  v = v.replace(/\\/g, "\\\\");
  v = v.replace(/\"/g, "\\\"");
  v = v.replace(/\n/g, "\\n");
  v = v.replace(/\t/g, "\\t");
  return v;
}

/* ~(MODULES::INC::VALIDATE, f'ControlObj::Validate
  <PRE>bool Validate()</PRE>
  This method validates the control according to the rules given in `Validator::Add' and returns
  true if the validation succeeds and false otherwise.
)~ */
function _VALCTL_Validate()
{ var arr = this.Data, v, len, fun, regex, err=false, msg;

  this.Value = (this.Value==null ? "" : String(this.Value));
  this.Typed = null;

  if((v=arr[_VAL_CLEAN]))
    this.Value = this.Value.replace(new RegExp(v, "gm"), "");
  if((v=arr[_VAL_EXECBV]))
  { fun = new Function("ctl", v);
    ret = fun(this);
    if(typeof(ret)=="string") err=true, msg=ret;
    else if(ret==false) return true;
  }
  if(!err && (v=arr[_VAL_DTYPE]) && this.Value>"" && v != "string" && _VAL_CoerceType(this.Value, v)==null)
    err=true, msg="is not a valid "+v+".\n\nIts value is:\n"+this.Value;
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
  if((v=arr[_VAL_EXECAV]))
  { fun = new Function("ctl", "err", "msg", v);
    ret = fun(this, err, msg);
    if(typeof(ret)=="string") err=true, msg=ret;
    else if(ret==false) return true;
  }
  if(err)
  { v = this.Data[_VAL_DESC];
    this.Parent.Message = (v ? v : this.Name)+' '+msg;
    return false;
  }
  return true;
}

/* ~(MODULES::INC::VALIDATE, f'ControlObj::HTML
  <PRE>string HTML([attr][, value])</PRE>
  This method returns the HTML for the control as described in `Validator::Write'.
)~ */
function _VALCTL_HTML(attr, value)
{ var type = this.Data[_VAL_TYPE], ret;
  this.Written = true;

  if(type==null) type="text";
  if(type=="textarea") ret="<textarea";
  else if(type=="select") ret="<select";
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
  { if(type=="radio")
    { if(value != null && this.Value==String(value)) ret += " checked=\"1\"";
    }
    else if(value != null) { value=this.Value=String(value); this.Typed=null; }
    else value=this.Value;

    if(type=="checkbox")
      ret += (value==null || value=="" || value=="0" || String(value).toLowerCase()=="false" ? " value=\"1\"" : " checked value=\"1\"");
    else if(value != null)
    { if(type=="textarea")
      { ret += '>'; ret += value; ret += "</textarea>";
        return ret;
      }
      else if(type!="select")
      { ret += " value=\""; ret += value.replace(/\"/g, "&quot;"); ret += "\"/>";
        return ret;
      }
    }
  }

  if(type=="textarea") return ret += "></textarea>";
  else if(type=="select") return ret += ">";
  else return ret += "/>";
}

/* ~(MODULES::INC::VALIDATE, f'ControlObj::Field
  <PRE>string Field()</PRE>
  This method returns the string value from the form data, cleaned if necessary.
  If the field was not submitted, an empty string is returned.
)~ */
function _VALCTL_Field()
{ var v=this.Parent.GetRawField(this.Name);
  if(v != null)
  { v = String(v);
    if(this.Data[_VAL_CLEAN])
    { var regex = new RegExp(this.Data[_VAL_CLEAN], "gm");
      v=v.replace(regex, "");
    }
    return v;
  }
  else return "";
}

/* ~(MODULES::INC::VALIDATE, f'ControlObj::SetVal
  <PRE>void SetVal(v)</PRE>
  This method sets the control's value properly. If v is null, the control's
  value is also set to null. Otherwise, it is set to the string representation
  of the value.
)~ */
function _VALCTL_SetVal(v)
{ this.Value=(v==null ? null : String(v));
  this.Typed=null;
}

function _VAL_CoerceType(val, type)
{ if(type==null || type=="string") return val;
  if(type=="int")
  { val=parseInt(val);
    return isNaN(val) ? null : val;
  }
  if(type=="float")
  { val=parseFloat(val);
    return isNaN(val) ? null : val;
  }
  if(type=="date")
  { try { val=new Date(val); } catch(e) { return null; }
    return val;
  }
  _VAL_Die("Unknown type "+type);
}

function _VAL_Die(s)
{ Response.Write("ERROR: ");
  Response.Write(s);
  throw new Error(-1, s);
}
</script>
