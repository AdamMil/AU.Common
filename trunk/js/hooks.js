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
var _HOOK_ONLOAD=0, _HOOK_ONUNLOAD=1, _HOOK_ONBEFOREUNLOAD=2;
var _HOOK_aHooks = new Array(null, null, null), _HOOK_aDone = new Array(false, false, false);

function _HOOK_Process(i)
{ _HOOK_aDone[i]=true;
  var i, f, a = _HOOK_aHooks[i];
  if(!a) return;
  for(i=0; i<a.length; i++)
  { f=a[i];
    if(typeof(f)=="function") f();
    else a[i][0](a[i][1]);
  }
}

function _HOOK_Add(i, f, p)
{ if(_HOOK_aDone[i]) f(p);
  else
  { var a=_HOOK_aHooks[i];
    if(a==null) a=_HOOK_aHooks[i]=new Array();
    a[a.length] = typeof(p)=="undefined" ? f : new Array(f,p);
  }
}

function HOOK_onLoad()
{ _HOOK_Process(_HOOK_ONLOAD);
  if(typeof(onLoad)=="function") onLoad();
}

function HOOK_onUnload()
{ if(typeof(onUnload)=="function") onUnload();
  _HOOK_Process(_HOOK_ONUNLOAD);
}

function HOOK_onBeforeUnload()
{ if(typeof(onBeforeUnload)=="function") onBeforeUnload();
  _HOOK_Process(_HOOK_ONBEFOREUNLOAD);
}

/* ~(MODULES::JS::HOOKS, t'hooks.js
  <PRE>
  function HOOK_onLoad();
  function HOOK_onUnload();
  function HOOK_onBeforeUnload();</PRE>

  function HOOK_addOnLoad(f[, p]);
  function HOOK_addOnUnload(f[, p]);
  function HOOK_addOnBeforeUnload(f[, p]);</PRE>

  The hooks.js file allows you to add a chain of hooks for certain events. The benefit
  of this is that multiple server-side ASP libraries can add javascript to the page
  that depends upon onload or onunload functionality without interfering with each
  other. Other events would be trivial to add to this code. The usage is that for
  events you want to handle, you add the HOOK_on*() functions as events on the BODY
  element:
  <PRE><body onload="HOOK_onLoad()" onbeforeunload="HOOK_onBeforeUnload()"></PRE>
  Then, to add a hook, just call HOOK_add*(), passing a function and an optional
  parameter. The function should be the function itself, not the name as a string.
  For instance:
  <PRE>HOOK_addOnLoad(my_func, its_param)</PRE>
  Finally, each event supports calling a function on the page if it exists. For
  instance, the HOOK_onLoad() function will call "onLoad()" if the function has
  been defined. This allows you to add an onLoad() function directly to the page
  which will be called by HOOK_onLoad(). For HOOK_onLoad(), the "onLoad" is
  called after the installed hooks. For HOOK_onUnload() and HOOK_onBeforeUnload(),
  the page functions (named "onUnload" and "onBeforeUnload") are called before
  the hooks.
)~ */
function HOOK_addOnLoad(f, p) { _HOOK_Add(_HOOK_ONLOAD, f, p); }
function HOOK_addOnUnload(f, p) { _HOOK_Add(_HOOK_ONUNLOAD, f, p); }
function HOOK_addOnBeforeUnload(f, p) { _HOOK_Add(_HOOK_ONBEFOREUNLOAD, f, p); }