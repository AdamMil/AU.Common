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

var _VAL_CurArr=null;
var _VAL_FieldList={"name":0, "desc":1, "maxlength":2, "minlen":3, "dtype":4, "clean":5, "nomatch":6,
                    "nmmsg":7, "match":8, "mmsg":9, "range":10, "execbv":11, "execbv":12};

// checks each element in a form using the CheckField() function
function VAL_Validate(frm)
{	var i, j, name, el=frm.elements, arr=_VAL_Forms[frm.id], len=el.length, jlen=arr.length;

	for(i=0; i<len; i++)
	{	name=el[i].name;
		for(j=0; j<jlen; j++)
		{	if(arr[j][0]==name)
			{	_VAL_CurArr=arr[j];
			  if(!VAL_ValidateField(el[i])) { _VAL_CurArr=null; return false; }
				break;
			}
		}
	}
	VAL_CurArr=null;
	return true;
}

function VAL_Disabled(fld)
{ return false;
}

function VAL_Visible(fld)
{ return true;
}

function VAL_EnableAll(cont, disable)
{
}

function VAL_Enable(ctl, disable)
{
}

function VAL_CmbSelect(fld, val)
{ var i, o=fld.options, len=o.length;
	for(i=0; i<len; i++) if(o[i].value==val) { o[i].checked=true; break; }
}

function _VAL_GetAttr(fld, attr)
{ var arr = null;
  if(_VAL_CurArr) arr = _VAL_CurArr;
  else
  { var i, len, name=fld.name;
    arr=_VAL_Forms[fld.form.id], len=arr.length;
    for(i=0; i<len; i++) if(arr[i][0]==name) { arr=arr[i]; break; }
    if(i==len) return null;
  }
  return arr[_VAL_FieldList[attr]];
}
