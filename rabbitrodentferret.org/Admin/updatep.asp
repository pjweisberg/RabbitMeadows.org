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
<center><font face=arial><h1>Manage Products</h1></font></center><hr>
<form method="post" action="addproduct.asp">
<font face=arial size=3><b> Add a new Product </b><input type="submit" name="newprod" value="Go"></font>
<hr>
</form>
<form method="post" action="addproduct.asp">
<font face=arial size=3><b>Edit or Delete an Existing Product</b>
<select name="Edprod" length=20>
<%
Dim objConn, strSQL
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = strconnect


strSQL = "SELECT ID, Name, Variation1Des, Variation1Price, Catagory FROM Products"  &_
                         " Order by Name"
objConn.Open
Dim objRS
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open strSQL, objConn

Do While not objRS.EOF



%>
<option value="<%=objRS("ID")%>"><%=objRS("ID")%> &nbsp; <%=objRS("Name")%>
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

