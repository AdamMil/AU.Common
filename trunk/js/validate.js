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

function VAL_Trim(s)
{ return s.replace(/^\s+|\s+$/g, "");
}

function VAL_TrimField(fld)
{ fld.value = VAL_Trim(fld.value);
}

function VAL_ValidPhone(fld, desc)
{ var val=VAL_Trim(fld.value), num=str.replace(/\D/g, ""), len=num.length;
  if(len<7 || (len>7 && len<10) || val.search(/[^0-9()\- ]/) != -1)
  { if(!desc) desc=_VAL_GetAttr(fld, "desc");
    alert((desc ? desc : "This")+" is not a valid phone number.");
    fld.focus();
    return false;
  }
  return true;
}

function VAL_FormatPhone(fld, desc)
{ VAL_TrimField(fld);
  if(fld.value.length==0 || !VAL_ValidPhone(fld, desc)) return false;
  var num=fld.value.replace(/\D/g, ""), len=num.length, st;
  if(len>10 && num.charAt(0) == '1') num=num.substr(1), len--;
  str  = len>7 ? '('+num.substr(0,st=3)+')' : "";
  str += num.substr(st,3) + '-' + num.substr(st+3,4);
  if(len>10) str+=' '+num.substr(st+7);
  fld.value=str;
  return true;
}

function VAL_ValidEmail(fld, desc, req)
{ var val=VAL_Trim(fld.value);
  if(req)
  { if(VAL_Missing(fld, desc)) return false;
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

function VAL_MatchingPW(fld1, fld2, desc)
{ if(!desc) desc=_VAL_GetDesc(fld1);
  if(desc) { desc=desc.split(' ')[0]; }
  if(!fld1.value.length || !fld2.value.length) alert((desc ? desc+" p" : "P")+"asswords are required fields!");
  else if(fld1.value != fld2.value) alert((desc ? desc+" p" : "P")+"asswords do not match.");
  else return true;
  fld1.focus();
  return false;
}

function VAL_ValidInt(fld, desc)
{ var val=VAL_Trim(fld.value), len=val.length;
  if(val.search(/\D/) != -1)
  { if(!desc) desc=_VAL_GetAttr(fld, "desc");
    alert((desc ? desc : "This")+" is not a valid integer.");
    fld.focus();
    return false;
  }
  return true;
}

function VAL_FormatInt(fld)
{ fld.value = fld.value.replace(/\D/g, "");
  if(fld.value=="") fld.value="0";
}

function VAL_ValidFloat(fld, desc)
{ var val=VAL_Trim(fld.value), pos;
  if(!/^\d*(\.\d+)?$/.test(val))
  { if(!desc) desc=_VAL_GetAttr(fld, "desc");
    alert((desc ? desc : "This")+" is not a valid number.");
    fld.focus();
    return false;
  }
  return true;
}

function VAL_FormatFloat(fld, desc)
{ fld.value=fld.value.replace(/[^\d.]/g, "");
  if(fld.value.length==0) { fld.value="0"; return; }
  var arr = fld.value.split(".");
  fld.value = arr[0] + (arr[1]>"" ? "."+arr[1] : "");
  VAL_ValidFloat(fld, desc);
}

function VAL_Missing(fld, desc, minlen)
{ if(VAL_Disabled(fld) || !VAL_Visible(fld)) return true;
  if(desc) desc=_VAL_GetAttr(fld, "desc");
  if(fld.value.length==0)
    alert((desc ? desc : "This") + " is a required field!");
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
    ret = fun(this);
    if(typeof(ret)=="string") err=true, msg=ret;
    else if(ret==false) return true;
  }
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
  if(err)
  { v = _VAL_GetAttr(fld, "desc");
    alert((v ? v : fld.name)+' '+msg);
    fld.focus();
    return false;
  }
  else if((v=_VAL_GetAttr(fld, "execav"))>"")
  { fun = new Function("ctl", v);
    fun(this);
  }
  return true;
}

// blurs all elements in a form before the form is submitted, to fix a bug
// in some browsers where onblur() and/or onchange() are not called when
// a form is submitted by pressing enter
function VAL_SubmitBlur(frm)
{ var i, e, el=frm.elements, len=el.length;
  for(i=0; i<len; i++)
  { e=el[i];
    if(e.onblur) e.onblur();
  }
}

// checks the radio button with a value equal to 'val'. it is passed an array
// with the radio button objects in it.. normally by just passing the name of
// the radio buttons
function VAL_RadCheck(fld, val)
{ for(var i=0; i<fld.length; i++)
  { if(fld[i].value==val)
    { fld[i].checked=true;
      return true;
    }
  }
  return false;
}

// returns the value of the checked radio button in an array of radio buttons,
// or null if none are checked.
function VAL_RadVal(fld)
{ for(var i=0; i<fld.length; i++) if(fld[i].checked) return fld[i].value;
  return null;
}

function VAL_SelVal(sel)
{ var i=sel.selectedIndex;
  return i==-1 ? null : sel.options[i].value;
}
