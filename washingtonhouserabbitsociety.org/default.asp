<% @language=vbscript %>
<% option explicit %>
<% Response.Buffer="true" %>
<!--#include file="strconn.asp"-->


<%
Dim adminpassword

adminpassword = Request.Form("adminpassword")

IF adminpassword <> "" then

    dim sqlStringP, objConn, objRSP

    sqlStringP = "SELECT AdminID FROM rebeccad.Admin " &_
    "WHERE adminPassword='" & adminpassword & "'" 

    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.ConnectionString=strconnect
    objConn.Open 
    SET objRSP = Server.CreateObject("ADODB.Recordset")
    objRSP.Open sqlStringP, objConn

      IF objRSP.EOF THEN 
      Response.Write "<font face=arial>You have entered an invalid password.  <br>"
      Response.Write "Please try again or obtain authorization from the site administrator.</font><br>"

      	objRSP.Close
	Set objRSP=nothing
	objConn.Close
	set objConn=nothing
%>
	<html>
	<body>
	<form method="post" action="default.asp">
	<font face="arial" size="5">Admin Password </font> <input type="text" name="adminpassword"><br>
	<input type = "submit" value="Log in">
	</form>
	</body>
	</html>     
     <%
        Else
        Response.Cookies("admin")="valid"
        Response.Cookies("admin").Expires=date+1
        Response.redirect "updatep.asp"
       end if
     %>
<% 
else

%>

<html>
<body>
<form method="post" action="default.asp">

<font face="arial" size="5">Admin Password </font> <input type="text" name="adminpassword"><br>
<input type = "submit" value="Log in">
</form>
</body>
</html>    
 
<%     
END IF
%>
    

  
																																																																																	