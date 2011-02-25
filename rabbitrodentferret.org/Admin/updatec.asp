<%@ LANGUAGE=VBSCRIPT %>
<% OPTION EXPLICIT %>
<% Response.Buffer="true" %>
<!--#include file="strconn.asp"-->

<%
If not Session("UserLevel")=2 then
	response.redirect("default.asp")

else
%>

<html>
<body>
<center><font face=arial><h1>Manage Featured Companions</h1></font></center><hr>
<form method="post" action="addcomp.asp">
<font face=arial size=3><b> Add a new Companion </b><input type="submit" name="newcomp" value="Go"></font>
<hr>
</form>
<form method="post" action="addcomp.asp">
<font face=arial size=3><b>Edit or Delete Rabbit Record</b>
<select name="Edcomp" length=20>
<%
Dim objConn, strSQL
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = strconnect

strSQL = "SELECT Photo, FirstName, Rotation, [Current], Owner FROM Rabbits" &_
                         " Order by Rotation Asc"
objConn.Open
Dim objRS
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open strSQL, objConn

Do While not objRS.EOF



%>
<option value="<%=objRS("Photo")%>"><b><%=objRS("Current")%></b> &nbsp; <%=objRS("FirstName")%> &nbsp;
<%=objRS("Rotation")%> &nbsp; <%=objRS("Owner")%> 
<%
objRS.MoveNext
Loop
objRS.close
set objRS=nothing
objConn.close
set objConn=nothing

%>
</select>
<input type="submit" value = "Go">
<input type="hidden" name="tablename" value=rabbits>
<hr>
</form>

<form method="post" action="addcomp.asp">
<font face=arial size=3><b>Edit or Delete Rodent Record</b>
<select name="Edcomp" length=20>
<%

Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = strconnect

strSQL = "SELECT Photo, FirstName, Rotation, [Current], Owner FROM Rodents" &_
                         " Order by Rotation Asc"
objConn.Open

Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open strSQL, objConn

Do While not objRS.EOF



%>
<option value="<%=objRS("Photo")%>"><b><%=objRS("Current")%></b> &nbsp; <%=objRS("FirstName")%> &nbsp;
<%=objRS("Rotation")%> &nbsp; <%=objRS("Owner")%> 
<%
objRS.MoveNext
Loop
objRS.close
set objRS=nothing
objConn.close
set objConn=nothing

%>
</select>
<input type="submit" value = "Go">
<input type="hidden" name="tablename" value=rodents>
<hr>
</form>

<form method="post" action="addcomp.asp">
<font face=arial size=3><b>Edit or Delete Ferret Record</b>
<select name="Edcomp" length=20>
<%

Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = strconnect

strSQL = "SELECT Photo, FirstName, Rotation, [Current], Owner FROM Ferrets" &_
                         " Order by Rotation Asc"
objConn.Open

Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open strSQL, objConn

Do While not objRS.EOF



%>
<option value="<%=objRS("Photo")%>"><b><%=objRS("Current")%></b> &nbsp; <%=objRS("FirstName")%> &nbsp;
<%=objRS("Rotation")%> &nbsp; <%=objRS("Owner")%> 
<%
objRS.MoveNext
Loop
objRS.close
set objRS=nothing
objConn.close
set objConn=nothing

%>
</select>
<input type="submit" value = "Go">
<input type="hidden" name="tablename" value=ferrets>
<hr>
</form>
<form action="default.asp" method="post">
	<input type="Submit" value="Back to the main Admin Page">
	</form>
</body>
</html>
<%
end if
%>

