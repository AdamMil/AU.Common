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

function VAL_Validate(frm)
{ var i, el=frm.elements, len=el.length;
  for(i=0; i<len; i++) if(!VAL_ValidateField(el[i])) return false;
  return true;
}

function VAL_Disabled(fld)
{ return fld.disabled;
}

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

function VAL_EnableAll(cont, bDisable)
{ var i, coll = cont.all.tags("INPUT");
  for(i=0; i<coll.length; i++) VAL_Enable(coll[i], disable);
  coll = cont.all.tags("TEXTAREA");
  for(i=0; i<coll.length; i++) VAL_Enable(coll[i], disable);
  coll = cont.all.tags("SELECT");
  for(i=0; i<coll.length; i++) VAL_Enable(coll[i], disable);
}

function VAL_Enable(ctl, bDisable)
{ ctl.disabled = bDisable ? true : false;
}

function VAL_CmbSelect(fld, val)
{ fld.value=val;
}

function _VAL_GetAttr(fld, attr)
{ var v=fld.getAttribute(attr);
  return v>"" ? v : null;
}