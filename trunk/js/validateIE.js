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

/* ~(MODULES::JS::VALIDATE, f'VAL_Validate
  <PRE>bool VAL_Validate(frm);</PRE>
  This method takes a form in 'frm' and validates the elements within it. This
  method must be used in conjunction with the `INC::VALIDATE::validate.inc' library.
  There must have been a `INC::VALIDATE::Validator' created for this form and
  `INC::VALIDATE::Validator::WriteCSCode' called. This method returns true if
  the validation succeeded and false otherwise (after alert() focusing on the
  field that failed).
)~ */
function VAL_Validate(frm)
{ var i, el=frm.elements, len=el.length;
  for(i=0; i<len; i++) if(!VAL_ValidateField(el[i])) return false;
  return true;
}

/* ~(MODULES::JS::VALIDATE, f'VAL_Disable
  <PRE>bool VAL_Disabled(fld);</PRE>
  This method takes a field in 'fld' and returns true if the field has been
  disabled. Note that in the validateNS.js implementation, this always returns
  false.
)~ */
function VAL_Disabled(fld)
{ return fld.disabled;
}

/* ~(MODULES::JS::VALIDATE, f'VAL_Visible
  <PRE>bool VAL_Visible(fld);</PRE>
  This method takes a field in 'fld' and returns true if the field is invisible
  (either because it or an ancestor node has a display:none or visibity:hidden
  style applied to it). Note that in the validateNS.js implementation, this
  always returns true.
)~ */
function VAL_Visible(fld)
{ while(fld)
  { try
    { if(fld.currentStyle.display=="none" || fld.currentStyle.visibility=="hidden") return false;
    }
    catch(e) { break; }
    fld = fld.parentNode;
  }
  return true;
}

/* ~(MODULES::JS::VALIDATE, f'VAL_EnableAll
  <PRE>void VAL_EnableAll(cont, bDisable);</PRE>
  This method takes a container (like a form) in 'cont' and either disables or
  enables all INPUT, TEXTAREA, and SELECT elements within it. If bDisabled is
  true, the fields are disabled. Otherwise, they are enabled.
)~ */
function VAL_EnableAll(cont, bDisable)
{ var i, coll = cont.all.tags("INPUT");
  for(i=0; i<coll.length; i++) VAL_Enable(coll[i], bDisable);
  coll = cont.all.tags("TEXTAREA");
  for(i=0; i<coll.length; i++) VAL_Enable(coll[i], bDisable);
  coll = cont.all.tags("SELECT");
  for(i=0; i<coll.length; i++) VAL_Enable(coll[i], bDisable);
}

/* ~(MODULES::JS::VALIDATE, f'VAL_Enable
  <PRE>void VAL_Enable(ctl, bDisable);</PRE>
  This method takes a field in 'fld' and either disables it.
  If bDisabled is true, the fields is disabled. Otherwise, it is enabled.
)~ */
function VAL_Enable(ctl, bDisable)
{ ctl.disabled = bDisable ? true : false;
}

/* ~(MODULES::JS::VALIDATE, f'VAL_CmbSelect
  <PRE>void VAL_CmbSelect(sel, val);</PRE>
  This method takes a selection box in 'sel' and changes the selection to the
  item that has a value matching 'val'. If 'sel' allows multiple selections,
  val should be a comma-delimited list of values to be selected.
)~ */
function VAL_CmbSelect(sel, val)
{ if(sel.multiple) VAL_CmbSelectMulti(sel, val);
  else fld.value=val;
}

function VAL_CmbSelectMulti(sel, val)
{ var vals = new Array();
  for(var v in val.split(",")) vals[v]=true;

  var i, o=fld.options, len=o.length;
	for(i=0; i<len; i++) o[i].checked = (vals[o[i].value]==true);
}

function _VAL_GetAttr(fld, attr)
{ var v=fld.getAttribute(attr);
  return v>"" ? v : null;
}
