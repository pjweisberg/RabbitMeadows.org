<% @language=vbscript %>
<% option explicit %>
<!--#include file="strconn.asp"-->

<%
If not Session("UserLevel")=2 then
	response.redirect("default.asp")

else

dim update
update=tRIM(Request.form("Edcomp"))
if update <> "" then
Dim objConn, strSQL, tab	
  	If Request("tablename")="rabbits" then
	tab="rabbits"
	strSQL="SELECT * FROM rebeccad.Rabbits WHERE Photo = '" & update & "'" 

  	Elseif REquest ("tablename")="rodents" then
	strSQL = "Select * from rebeccad.Rodents where Photo = '" & update & "'"
	tab="rodents"

  	Elseif Request("tablename")="ferrets" then
	strSql = "SELECT * from rebeccad.ferrets where Photo = '" & update & "'"
	tab="ferrets"
	
	End if


Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = strconnect
objConn.Open

Dim objRS
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open strSQL, objConn

dim edPhoto, edWidth, edHeight
dim edFirstName, edCaption, edRotation, edDisplay
dim edOwner, edOwneremail

Do while not objRS.EOF
edPhoto= Trim(objRS("Photo"))
edFirstName=objRS("FirstName")
If not isNull(edFirstName) then
edFirstName = Server.HtmlEncode(edFirstName)
end if
edCaption = objRS("Caption")
edWidth = objRS("Width")
edHeight = objRS("Height")
edRotation= objRS("Rotation")
edDisplay = objRS ("Current")
if edDisplay<>1 then
edDisplay=0
end if

edOwner = objRS("Owner")
if edOwner="" then
edOwner=Null
end if

IF tab="rabbits" or tab="rodents" then
edOwneremail= objRS("Owner e-mail")
elseif tab="ferrets" then
edOwneremail=objRS("Owner email")
end if

if edOwneremail="" then
edOwneremail=Null
end if


objRS.Movenext
Loop

objRS.Close
set objRS = nothing
objConn.Close
set objConn=nothing

end if
%>
<html>
<head>
<script language="javascript">

function preview() {
 
DispWin = window.open('','Newwin','toolbar=no,status=no,width=500,height=300')

 
 	message="<html><body><table BORDER=0 width=446 bgcolor=#BDD5EB><tr><TD valign=center align=center>";     
        message+= "<img width=" + document.newcomp.width.value;
        message+= " height = " + document.newcomp.height.value + " src=http://www.rabbithouse.org/features/" + document.newcomp.image.value;
	message+= " vspace=10 hspace=10 align=right>";
   	message+= "<p align=middle><font color=#000000 face=arial><br><h3><b>" + document.newcomp.name.value + "</b></h3></p>";
	message+= "<p align=left valign=middle>" + document.newcomp.caption.value + "</p></font></td></tr></table>"       
        
    DispWin.document.write(message);
}
</script>
</head>
<body>

<form name="newcomp" method="post" action="addcomp2.asp">
<center><font face=arial size=5>

<%
if update <> "" then
Response.Write "Edit a Companion"
else
Response.Write "Add a Companion"
end if
%>
</font><hr width=40% size=1 noshade></center>
<font face=arial>Name or Title <input type="text" name = "name" value="<%=edFirstName%>"><br>
Caption <br> <textarea name="caption" size="45" rows=5 cols=40><%=edCaption%></textarea><br>
Image filename <input type="text" name="image" value="<%=edPhoto%>"><br>
Image width <input type="text" name="width" value="<%=edWidth%>"><br>
Image height <input type="text" name="height" value="<%=edHeight%>"><br>

Table <select name="tablename">
<%
dim selected
selected=""
if tab="rabbits" then
selected="selected" 
end if
response.write "<option value=rabbits " & selected & ">Rabbits"

selected = ""
if tab="rodents" then
selected="selected" 
end if
response.write "<option value=rodents " & selected & ">Rodents"

selected=""
if tab="ferrets" then
selected="selected" 
end if
response.write "<option value=ferrets " & selected & ">Ferrets"

%>
</select><br>
<%

%>
Rotation <select name="rotation">
<%

dim i, catname
for i=1 to 9
selected=""
if edRotation=i then
selected="selected"
end if

 Select Case i

 Case "1"
  catname="Toys"
 Case "2"
  catname="Housing"
 Case "3"
  catname="Furnishings"
 Case "4"
  catname="Food"
 Case "5"
  catname="Grooming"
 Case "6"
  catname="Healthcare"
 Case "7"
  catname="People Stuff"
 Case "8"
  catname="Books & Videos"
 Case "9"
  catname = "Misc"
 End Select

REsponse.Write "<option value=" & i & " " & selected & ">" & catname

next

%>
</select><br>
<%
if edDisplay=0 then
%>
Display <input type="checkbox" value="1" name="display" unchecked><br>
<%
else
%>

Display <input type="checkbox" value="1" name="display" checked><br>
<%
end if
%>

Owner Name <input type="text" name="owner" size=30 value="<%=edOwner%>"><br>
Owner E-mail <input type="text" name="owneremail" size=50 value="<%=edOwneremail%>">
<br><input type="submit" value="Submit">
<input type="hidden"  name="Edcomp" value="<%=update%>">
<input type="button" value="Preview" onClick="preview();">
</form>
<form method="post" action="addcomp2.asp">
<input type="submit" value="Delete this product">
<input type="hidden" name = "delete" value="delete">
<input type="hidden" name="Edcomp" value="<%=update%>">
<input type="hidden" name="tablename" value="<%=tab%>">
</form>
<form action="updatec.asp">
<input type="submit" value="Cancel">
</form>
</body>
</html>
<%
end if
%>
