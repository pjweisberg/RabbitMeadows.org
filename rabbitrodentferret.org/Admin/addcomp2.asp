<% @language=vbscript %>
<% Option Explicit %>
<!--#include file = "adovbs.inc"-->
<!--#include file="strconn.asp"-->

<% Response.Buffer = "True" %>


<%
If not Session("UserLevel")=2 then
	response.redirect("default.asp")

else

dim Name, Caption, Image, Width, Height
Dim Tablename, Rotation, Display, Owner, Owneremail

Name = Request("name")
Caption = Request("caption")
Image = Request("image")
Width=Request("width")
If Width="" then 
Width = Null
end if
Height = Request ("height")
If Height="" then
Height = Null
end if

Tablename = "rebeccad." & Request ("tablename")

Rotation=Request("rotation")
Display=Request("display")
if Display <>1 then
Display=0
end if
Owner=Request("owner")
if Owner="" then
Owner=Null
end if

Owneremail=Request("owneremail")
if Owneremail="" then
Owneremail=Null
end if



Dim EdID, Delete
EdID=Request.Form("EdComp")
Delete = REquest.Form("delete")
If Delete <> "" then
response.write Tablename & "<br>" & EdID

dim objConnD, objRSD, sqlStringD
set objConnD=Server.CreateObject("ADODB.Connection")
objConnD.ConnectionString=strconnect
objConnD.Open
Set objRSD = Server.CreateObject("ADODB.Recordset")
sqlStringD = "Select * From " & Tablename & " where Photo ='" & EdID & "'"
objRSD.ActiveConnection = objConnD
objRSD.LockType=adLockOptimistic
objRSD.Open sqlStringD
objRSD.Delete
objRSD.Update
objRSD.Close
set objRSD=nothing
objConnD.Close
Set objConnD=Nothing

Else
  If EdID<>"" then

  Dim objConn, sqlString
  set objConn = Server.CreateObject("ADODB.Connection")
  objConn.ConnectionString=strconnect
  objConn.Open
  dim objRS

  set objRS = Server.CreateObject("ADODB.Recordset")
  sqlString = "Select * From " & Tablename & " Where Photo= '" & EdID & "'"
  objRS.ActiveConnection = objConn
  objRS.lockType=adLockOptimistic
  objRS.Open sqlString

  objRS("FirstName")=Name
   objRS("Caption") = Caption
  objRS("Width") = Width
  objRS("Height") = Height
  objRS("Rotation") = Rotation
  objRS("Current") = Display
  objRS("Owner") = Owner

If Tablename="rebeccad.Rabbits" or Tablename="rebeccad.Rodents" then
  objRS("Owner e-mail") = Owneremail
elseif Tablename="rebeccad.Ferrets" then
objRS("Owner email") = Owneremail
end if

    objRS.Update
  objRS.Close
  set objRS=nothing
  objConn.Close
  set objConn=nothing

  Else
  Dim objConn2, objRS2
  set objConn2 = Server.CreateObject("ADODB.Connection")
  objConn2.ConnectionString=strconnect
  objConn2.Open

  set objRS2 = Server.CreateObject("ADODB.Recordset")
  objRS2.Open Tablename, objConn2, , adLockOptimistic, adCmdTable


  objRS2.AddNew
  objRS2("FirstName")=Name
  objRS2("Photo")=Image
   objRS2("Caption") = Caption
  objRS2("Width") = Width
  objRS2("Height") = Height
  objRS2("Rotation") = Rotation
  objRS2("Current") = Display
  objRS2("Owner") = Owner

If Tablename="rebeccad.Rabbits" or Tablename="rebeccad.Rodents" then
objRS2("Owner e-mail") = Owneremail
elseif Tablename="rebeccad.Ferrets" then
  objRS2("Owner email") = Owneremail
end if

  objRS2.Update
  objRS2.Close
  set objRS2 = nothing
  objConn2.Close
  set objConn2 = nothing
  end if
end if
Response.redirect "updatec.asp"
end if

%>

