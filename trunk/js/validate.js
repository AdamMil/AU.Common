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

// calls functions from validateIE.js/validateNS.js

/* ~(MODULES::JS::VALIDATE, t'validate.js
  The validate.js file and the accompanying validateIE.js and validateNS.js
  files implement functions helpful for validation and form manipulation.
  This library was originally intended to be used by `INC::VALIDATE::validate.inc',
  but it can also be used by itself. If validate.js is included, the
  appropriate helper file validateIE.js or validateNS.js should be included.
  validateIE is designed for browser with a browser level of at least 2, while
  validateNS.js should work with a browser level of 1 (see `INC::BROWSER::browser.inc').
  A number of functions are named like VAL_Valid*() and VAL_Format*(). The
  VAL_Valid*() functions validate a field but do not change it substantially.
  If the field does not pass the validation, an alert() is done, the focus
  is moved to the field, and the function returns false. Otherwise, it returns true.
  The VAL_Format*() functions attempt to validate and also format the field.
  If the validation fails, the field may or may not have been changed, but in any case,
  if the format function returns true, the field contains a valid value. If it returns
  false, the field could not successfully be formatted. Most VAL_Valid*() and
  VAL_Format*() functions take a 'req' parameter. If req is false (the default), the
  validation/format will also succeed if the field is empty, disabled, or invisible.
  Note that VAL_Format*() functions will still format disabled/invisible fields.
)~ */

/* ~(MODULES::JS::VALIDATE, f'VAL_Trim
  <PRE>string VAL_Trim(s);</PRE>
  The VAL_Trim function trims the leading and trailing whitespace from the string
  passed as 's'. 's' must be a string or an exception will be thrown. The trimmed
  string is returned.
)~ */
function VAL_Trim(s)
{ return s.replace(/^\s+|\s+$/g, "");
}

/* ~(MODULES::JS::VALIDATE, f'VAL_TrimField
  <PRE>string VAL_TrimField(fld);</PRE>
  The VAL_TrimField takes a control in 'fld' and trims the fld.value property. The
  trimmed value is both stored in fld.value and returned.
)~ */
function VAL_TrimField(fld)
{ return fld.value = VAL_Trim(fld.value);
}

/* ~(MODULES::JS::VALIDATE, f'VAL_ValidPhone
  <PRE>bool VAL_ValidPhone(fld, desc, req);</PRE>
  This method takes a field in 'fld' and an optional field description.
  It then attempts to check that the field contains a valid phone number.
  See `validate.js' for more information.
)~ */
function VAL_ValidPhone(fld, desc, req)
{ var val=VAL_Trim(fld.value), num=str.replace(/\D/g, ""), len=num.length;
  if(!req && (len==0 || VAL_IsMissing(fld))) return true;
  if(len<7 || (len>7 && len<10) || val.search(/[^0-9()\- ]/) != -1)
  { if(!desc) desc=_VAL_GetAttr(fld, "desc");
    alert((desc ? desc : "This")+" is not a valid phone number.");
    fld.focus();
    return false;
  }
  return true;
}

/* ~(MODULES::JS::VALIDATE, f'VAL_FormatPhone
  <PRE>bool VAL_FormatPhone(fld, desc, req);</PRE>
  This method first calls `VAL_ValidPhone' to check the validity of the field.
  If the field is valid, the phone number is formatted to match the form
  "(AAA) PPP-SSSS EEEE". The area code and extension will be dropped if they
  are not in the input. See `validate.js' for more information.
)~ */
function VAL_FormatPhone(fld, desc, req)
{ VAL_TrimField(fld);
  if(!VAL_ValidPhone(fld, desc, req)) return false;
  var num=fld.value.replace(/\D/g, ""), len=num.length, st;
  if(len==0) return true;
  if(len>10 && num.charAt(0) == '1') num=num.substr(1), len--;
  str  = len>7 ? '('+num.substr(0,st=3)+')' : "";
  str += num.substr(st,3) + '-' + num.substr(st+3,4);
  if(len>10) str+=' '+num.substr(st+7);
  fld.value=str;
  return true;
}

/* ~(MODULES::JS::VALIDATE, f'VAL_ValidEmail
  <PRE>bool VAL_ValidEmail(fld, desc, req);</PRE>
  This method takes a field in 'fld' and an optional field description.
  It then attempts to check that the field contains a valid email address.
  See `validate.js' for more information.
)~ */
function VAL_ValidEmail(fld, desc, req)
{ var val=VAL_TrimField(fld);
  if(req)
  { if(!VAL_Require(fld, desc)) return false;
  }
  else if(val.length==0) return true;

  if(!/^([\w\.\-])+@[a-z\d\-]?(\.[a-z\d\-]+)+$/i.test(val))
  { if(!desc) desc=_VAL_GetAttr(fld, "desc");
    alert((desc ? desc : "This")+" is not a valid email address.");
    fld.focus();
    return false;
  }
  return true;
}

/* ~(MODULES::JS::VALIDATE, f'VAL_MatchingPW
  <PRE>bool VAL_MatchingPW(fld1, fld2, desc);</PRE>
  This method takes two password fields and makes sure that they match.
  If no description is passed, the description of the first field is used
  if possible. If the description contains more than one word, only the first
  word is used as the description. For instance, "Login Password" becomes
  "Login" for purposes of the error message, since "passwords" is the first
  word of the error message (ie "Login passwords do not match.").
  See `validate.js' for more information.
)~ */
function VAL_MatchingPW(fld1, fld2, desc, req)
{ if(!desc) desc=_VAL_GetAttr(fld1, "desc");
  if(desc) { desc=desc.split(' ')[0]; }
  if(req && (!fld1.value.length || !fld2.value.length)) alert((desc ? desc+" p" : "P")+"asswords are required fields!");
  else if(fld1.value != fld2.value) alert((desc ? desc+" p" : "P")+"asswords do not match.");
  else return true;
  fld1.focus();
  return false;
}

/* ~(MODULES::JS::VALIDATE, f'VAL_ValidInt
  <PRE>bool VAL_ValidInt(fld, desc, req);</PRE>
  This method takes a field in 'fld' and an optional field description.
  It then attempts to check that the field contains a valid integer.
  See `validate.js' for more information.
)~ */
function VAL_ValidInt(fld, desc, req)
{ var val=VAL_Trim(fld.value), len=val.length;
  if(!req && (val.length==0 || VAL_IsMissing(fld))) return true;
  if(val.length==0 || val.search(/\D/) != -1)
  { if(!desc) desc=_VAL_GetAttr(fld, "desc");
    alert((desc ? desc : "This")+" is not a valid integer.");
    fld.focus();
    return false;
  }
  return true;
}

/* ~(MODULES::JS::VALIDATE, f'VAL_FormatInt
  <PRE>bool VAL_FormatInt(fld);</PRE>
  This method takes a field in 'fld' and formats it into a valid integer.
  All non-digit characters are removed, and if there is nothing left, the
  field is set to 0.
  See `validate.js' for more information.
)~ */
function VAL_FormatInt(fld)
{ fld.value = fld.value.replace(/\D/g, "");
  if(fld.value.length==0) fld.value="0";
  return true;
}

/* ~(MODULES::JS::VALIDATE, f'VAL_ValidFloat
  <PRE>bool VAL_ValidFloat(fld, desc, req);</PRE>
  This method takes a field in 'fld' and an optional field description.
  It then attempts to check that the field contains a valid number.
  See `validate.js' for more information.
)~ */
function VAL_ValidFloat(fld, desc, req)
{ var val=VAL_Trim(fld.value), pos;
  if(!req && (val.length==0 || VAL_IsMissing(fld))) return true;
  if(val.length==0 || !/^\d*(\.\d+)?$/.test(val))
  { if(!desc) desc=_VAL_GetAttr(fld, "desc");
    alert((desc ? desc : "This")+" is not a valid number.");
    fld.focus();
    return false;
  }
  return true;
}

/* ~(MODULES::JS::VALIDATE, f'VAL_FormatFloat
  <PRE>bool VAL_FormatFloat(fld, desc, req);</PRE>
  This method takes a field in 'fld' and attempts to make a valid number out of it.
  `VAL_ValidFloat' is then called to check whether the attempt was successful.
  See `validate.js' for more information.
)~ */
function VAL_FormatFloat(fld, desc, req)
{ fld.value=fld.value.replace(/[^\d.]/g, "");
  if(fld.value.length==0) { fld.value="0"; return; }
  var arr = fld.value.split(".");
  fld.value = arr[0] + (arr[1]>"" ? "."+arr[1] : "");
  return VAL_ValidFloat(fld, desc, req);
}

/* ~(MODULES::JS::VALIDATE, f'VAL_ValidDate
  <PRE>bool VAL_ValidDate(fld, desc, req);</PRE>
  This method takes a field in 'fld' and an optional field description.
  It then attempts to check that the field contains a valid date by invoking the
  JScript date parser. If the date could not be parsed by JScript, it is assumed
  to be invalid.
  See `validate.js' for more information.
)~ */
function VAL_ValidDate(fld, desc, req)
{ var val=VAL_Trim(fld.value);
  if(!req && (val.length==0 || VAL_IsMissing(fld))) return true;
  try { var d = new Date(val); }
  catch(e)
  { if(!desc) desc=_VAL_GetAttr(fld, "desc");
    alert((desc ? desc : "This")+" is not a valid date.");
    fld.focus();
    return false;
  }
  return true;
}

/* ~(MODULES::JS::VALIDATE, f'VAL_FormatDate
  <PRE>bool VAL_FormatDate(fld, desc, req);</PRE>
  This method takes a field in 'fld'. If the date is valid according to `VAL_ValidDate',
  it is converted into "MM/DD/YYYY" format.
  See `validate.js' for more information.
)~ */
function VAL_FormatDate(fld, desc, req)
{ var val=VAL_TrimField(fld);
  if(!req && val.value.length==0) return true;
  if(!VAL_ValidDate(fld, desc, req)) return false;
  var d=new Date(val);
  fld.value = (d.getMonth()+1)+'/'+d.getDate()+'/'+d.getFullYear();
  return true;
}

/* ~(MODULES::JS::VALIDATE, f'VAL_IsMissing
  <PRE>bool VAL_IsMissing(fld);</PRE>
  This method returns true if the field passed is empty, disabled, or invisible.
)~ */
function VAL_IsMissing(fld)
{ return fld.value.length==0 || VAL_Disabled(fld) || !VAL_Visible(fld));
}

/* ~(MODULES::JS::VALIDATE, f'VAL_Require
  <PRE>bool VAL_Require(fld, desc, minlen);</PRE>
  This method takes a field in 'fld' and an optional field description. 'minlen',
  if passed, controls the minimum allowed length of the field. The field is not
  trimmed before being checked. If the field is disabled or invisible, the
  function returns true. This function behaves like a VAL_Valid*() function,
  returning true/false based on whether the field is filled out or not.
  See `validate.js' for more information.
)~ */
function VAL_Require(fld, desc, minlen)
{ if(VAL_Disabled(fld) || !VAL_Visible(fld)) return true;
  if(desc) desc=_VAL_GetAttr(fld, "desc");
  if(fld.value.length==0)
    alert((desc ? desc : "This")+" is a required field!");
  else if(minlen && fld.value.length<minlen)
    alert((desc ? desc : "Field")+" must be at least "+minlen+" characters.");
  else return true;
  fld.focus();
  return false;
}

function VAL_ValidateField(fld)
{ if(VAL_Disabled(fld) || !VAL_Visible(fld)) return true;
  var v, len, fun, regex, err=false, msg;

  if((v=_VAL_GetAttr(fld, "clean")))
    fld.value = fld.value.replace(new RegExp(v, "gm"), "");
  if((v=_VAL_GetAttr(fld, "execbv")))
  { fun = new Function("ctl", v);
    ret = fun(fld);
    if(typeof(ret)=="string") err=true, msg=ret;
    else if(ret==false) return true;
  }
  if(!err && (v=_VAL_GetAttr(fld, "dtype")) && _VAL_CoerceType(fld.value, v)==null)
    err=true, msg="cannot be converted to a"+v;
  if(!err && (v=_VAL_GetAttr(fld, "maxlength")) && fld.value.length>v)
    err=true, msg="must be at most "+v+" characters.";
  if(!err && (v=_VAL_GetAttr(fld, "minlen")))
  { len = fld.value.length;
    if(len < v)
      err=true, msg = (v==1 ? "is required." : "must be at least "+v+" characters.");
    else if(v<0 && len>0 && len<-v)
      err=true, msg = "must be at least "+(-v)+" characters, if specified.";
  }
  if(!err && (v=_VAL_GetAttr(fld, "nomatch")))
  { if(fld.value.search(new RegExp(v, "gm")) != -1)
    { err=true, msg=_VAL_GetAttr(fld, "nmmsg");
      if(!msg) msg="is invalid (matches "+v+").";
    }
  }
  if(!err && (v=_VAL_GetAttr(fld, "match")))
  { if(fld.value.search(new RegExp(v, "gm")) == -1)
    { err=true, msg=_VAL_GetAttr(fld, "mmsg");
      if(!msg) msg="fails to match required format ("+v+").";
    }
  }
  if(!err && (v=_VAL_GetAttr(fld, "range")))
  { var val = parseFloat(fld.value);
    v = v.split(' ');
    if(val<parseFloat(v[0]) || val>parseFloat(v[1]))
      err=true, msg = "is outside of expected range ("+v[0]+" to "+v[1]+").";
  }
  if((v=_VAL_GetAttr(fld, "execav")))
  { fun = new Function("ctl", "err", "msg", v);
    ret = fun(fld, err, msg);
    if(typeof(ret)=="string") err=true, msg=ret;
    else if(ret==false) return true;
  }
  if(err)
  { v = _VAL_GetAttr(fld, "desc");
    alert((v ? v : fld.name)+' '+msg);
    fld.focus();
    return false;
  }
  return true;
}

/* ~(MODULES::JS::VALIDATE, f'VAL_SubmitBlur
  <PRE>void VAL_SubmitBlur(frm);</PRE>
  This method takes a form in 'frm' and fires the onblur event of every element in the form.
  This is to work around an issue in some browsers in which onblur for the current field
  is not called if the form is submitted by pressing enter/return.
)~ */
function VAL_SubmitBlur(frm)
{ var i, e, el=frm.elements, len=el.length;
  for(i=0; i<len; i++)
  { e=el[i];
    if(e.onblur) e.onblur();
  }
}

/* ~(MODULES::JS::VALIDATE, f'VAL_RadCheck(fld, val)
  <PRE>bool VAL_RadCheck(fld, val);</PRE>
  This method takes an array of radio buttons in 'fld' and a value in 'val'. The radio
  button that has a value equal to 'val' will be checked. Normally, just passing the
  name of the radio button group will pass an array containing the radio buttons.
  This method returns true if the radio button was found and false otherwise.
)~ */
function VAL_RadCheck(fld, val)
{ for(var i=0; i<fld.length; i++)
  { if(fld[i].value==val)
    { fld[i].checked=true;
      return true;
    }
  }
  return false;
}

/* ~(MODULES::JS::VALIDATE, f'VAL_RadVal(fld)
  <PRE>variant VAL_RadVal(fld);</PRE>
  This method takes an array of radio buttons in 'fld' and returns the value of the
  radio button that is checked. If no radio button is checked, the method returns
  null. Normally, just passing the name of the radio button group will pass an array
  containing the radio buttons.
)~ */
function VAL_RadVal(fld)
{ for(var i=0; i<fld.length; i++) if(fld[i].checked) return fld[i].value;
  return null;
}

/* ~(MODULES::JS::VALIDATE, f'VAL_SelVal(sel)
  <PRE>variant VAL_SelVal(sel);</PRE>
  This method takes a selection box in 'sel' and returns the value of the selected item.
  If no item is selected, null is returned.
)~ */
function VAL_SelVal(sel)
{ var i=sel.selectedIndex;
  return i==-1 ? null : sel.options[i].value;
}

function _VAL_CoerceType(val, type) // assumes 'val' is a string
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
  throw new Error(0, "Unknown type "+type);
}

function _VAL_InitForm(frm)
{ WRITE ME
}
