<%
  var SessionID, SessionURL;
  if(!oSession("User.LoggedIn") && !oSession("User.LoggedOut"))
  { var sUser = String(Request.ServerVariables("AUTH_USER"));
    if(sUser>"")
    { var db=oDBMan.CreateDB(), rs=db.ExecuteO("sp_GetEmployeeInfo", "@LoginName", sUser.substr(sUser.lastIndexOf('\\')+1));
      if(!rs.EOF)
      { oSession("User.Login") = rs("LoginName").Value;
        oSession("User.Name")  = rs("FirstName")+" "+rs("LastName");
        oSession("User.LoggedIn") = true;
      }
      db = rs = null;
    }
  }
  SessionID  = oSession.SessionID;
  SessionURL = "Session="+Server.URLEncode(SessionID);
%>