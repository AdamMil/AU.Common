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

var UA_Agent="Unknown", UA_Browser=UA_Agent, UA_Maker=UA_Agent, UA_Version=0, UA_Level=0, UA_bLoaded=false;

/* ~(MODULES::INC::BROWSER, t'browser.inc
  <PRE>function UA_Init();</PRE>
  The browser module consists of a single function, UA_Init(). Calling this
  sets several global variables:
  <PRE>
  UA_Agent   - the complete user agent string
  UA_Browser - the name of the browser (ie Mozilla, MSIE, Opera)
  UA_Maker   - the maker of the browser (ie Netscape or Microsoft)
  UA_Version - the version of the browser
  UA_Level   - the "level" of the browser:
    0 - the browser offers no support for javascript
    1 - the browser offers support compatible with NS 4
    2 - the browser offers support compatible with IE 5
    3 - the browser offers support compatible with IE 6
  UA_bLoaded - set to true if UA_Init() has been called
  </PRE>
  The most useful of these is the UA_Level variable, which can be used to
  tailor output to the browser. Note, however, that the is code works
  entirely upon the user agent passed by the browser. This string can be
  forged (some browsers do this). In addition, the code does not detect
  whether javascript is actually enabled on the client.
*/

function UA_Init()
{ if(UA_bLoaded) return;

  UA_Agent   = String(Request.ServerVariables("HTTP_USER_AGENT"));
  UA_Browser = UA_Maker = "Unknown";
  UA_Version = 0;
  UA_Level   = 0;
  UA_bLoaded = true;
  
  // first check for IE
  var pos = UA_Agent.indexOf("MSIE");
  if(pos != -1)
  { UA_Browser="MSIE";
    UA_Maker  ="Microsoft";
    UA_Version=parseFloat(UA_Agent.substr(pos+5));
    UA_Level  =UA_Version >= 6 ? 3 : UA_Version >= 5.01 ? 2 : 1;
    return;
  }
  
  // if it's not a known browser, use this generic code
  var reg = new RegExp("^\\s*(\S*)[^(]*\\((.*)\\).*?$");
  var part1, part2;
  if(reg.exec(UA_Agent) != null) part1 = RegExp.$1, part2 = RegExp.$2;
  else
  { reg = new RegExp("^\\s*(\\S*).*$");
    if(reg.exec(UA_Agent) !=  null) part1 = RegExp.$1;
    else part1 = UA_Agent;
  }

  pos = part1.indexOf('/');
  if(pos != -1)
  { UA_Browser=part1.substr(0, pos);
    UA_Version=parseFloat(part1.substr(pos+1));
  }
  else UA_Browser=part1;
  
  if(UA_Browser=="Mozilla")
  { if(UA_Agent.indexOf("compatible")==-1) UA_Maker="Netscape";
    UA_Level = UA_Version>=4 ? 1 : 0;
  }
  else if(UA_Browser=="Opera") UA_Level = UA_Version>=5 ? 1 : 0;
}
</script>
