<% @language=vbscript %>
<% Option Explicit %>
<!--#include file = "adovbs.inc"-->
<!--#include file="strconn.asp"-->

<% Response.Buffer = "True" %>


<%
If not Session("UserLevel")=2 then
	response.redirect("default.asp")

else

dim Name, Image, Width, Height, Catagory, Rabbit, Rodent, Ferret, Description, Variations
Dim Status, Varlist, Pricelist, InOrder
Dim Vard1, Vard2, Vard3, Vard4, Vard5, Vard6, Vard7, Vard8, Vard9, Vard10
Dim Varp1, Varp2, Varp3, Varp4, Varp5, Varp6, Varp7, Varp8, Varp9, Varp10

Name = Request("name")

Image = Request("image")
If Image = "" then
Image = Null
else  Image = Image & " "" "
end if

Width=Request("width")
If Width="" then 
Width = Null
end if

Height = Request ("height")
If Height="" then
Height = Null
end if

Catagory = Request ("catagory")

Rabbit=Request("rabbits")
If Rabbit= "" then
Rabbit = 0
End if

Rodent=Request("rodents")
If Rodent = "" then
Rodent = 0
end if

Ferret=Request("ferrets")
If Ferret = "" then
Ferret = 0
end if

Description=REquest("descriptionlng")

Status = Request("status")
If Status <> 1 then
Status = 0
end if

Vard1 = Trim(Request("vard1"))
Vard2 = Trim(Request("vard2"))
Vard3 = Trim(Request("vard3"))
Vard4 = Trim(Request("vard4"))
Vard5 = Trim(Request("vard5"))
Vard6 = Trim(Request("vard6"))
Vard7 = Trim(Request("vard7"))
Vard8 = Trim(Request("vard8"))
Vard9 = Trim(Request("vard9"))
Vard10 = Trim(Request("vard10"))



If Vard1 <> "" and not isnull(Vard1) then
Varlist = Vard1
Variations = 1
end if

If Vard2 <> "" and not isnull(Vard2) then
Varlist = (Vard1 & ", " & Vard2)
Variations = 2
end if

If Vard3 <> "" and not isnull(Vard3) then
Varlist = (Vard1 & ", " & Vard2 & ", " & Vard3)
Variations = 3
end if

If Vard4<> "" and not isnull(Vard4) then

Varlist = (Vard1 & ", " & Vard2 & ", " & Vard3 & ", " & Vard4)
Variations=4
end if

If Vard5 <> "" and not isnull(Vard5) then
Varlist = (Vard1 & ", " & Vard2 & ", " & Vard3 & ", " & Vard4 & ", " & Vard5)
Variations = 5
end if

If Vard6 <> "" and not isnull(Vard6) then
Varlist = (Vard1 & ", " & Vard2 & ", " & Vard3 & ", " & Vard4 & ", " &_
 Vard5 & ", " & Vard6)
Variations = 6
end if

If Vard7 <> "" and not isnull(Vard7) then
Varlist = (Vard1 & ", " & Vard2 & ", " & Vard3 & ", " & Vard4 & ", " &_
 Vard5 & ", " & Vard6 & ", " & Vard7)
Variations = 7
end if

If Vard8 <> "" and not isnull(Vard8) then
Varlist = (Vard1 & ", " & Vard2 & ", " & Vard3 & ", " & Vard4 & ", " &_
 Vard5 & ", " & Vard6 & ", " & Vard7 & ", "  & Vard8)
Variations = 8
end if


If Vard9 <> "" and not isnull(Vard9) then
Varlist = (Vard1 & ", " & Vard2 & ", " & Vard3 & ", " & Vard4 & ", " &_
 Vard5 & ", " & Vard6 & ", " & Vard7 & ", "  & Vard8 & ", " & Vard9)
Variations = 9
end if

If Vard10 <> "" and not isnull(Vard10) then
Varlist = (Vard1 & ", " & Vard2 & ", " & Vard3 & ", " & Vard4 & ", " &_
 Vard5 & ", " & Vard6 & ", " & Vard7 & ", "  & Vard8 & ", " & Vard9 &_
", " & Vard10)
Variations = 10
end if

If Varlist = "" then
Varlist = Null
End if


'-------------------------------------------------

Varp1 = Trim(Request("varp1"))
Varp2 = Trim(Request("varp2"))
Varp3 = Trim(Request("varp3"))
Varp4 = Trim(Request("varp4"))
Varp5 = Trim(Request("varp5"))
Varp6 = Trim(Request("varp6"))
Varp7 = Trim(Request("varp7"))
Varp8 = Trim(Request("varp8"))
Varp9 = Trim(Request("varp9"))
Varp10 = Trim(Request("varp10"))


Pricelist = Varp1

If Varp2 <> "" and not isnull(Varp2)  then
Pricelist = (Varp1 & ", " & Varp2)
end if

If Varp3 <> "" and not isnull(Varp3) then
Pricelist = (Varp1 & ", " & Varp2 & ", " & Varp3)
end if

If Varp4<> "" and not isnull(Varp4) then
Pricelist = (Varp1 & ", " & Varp2 & ", " & Varp3 & ", " & Varp4)
end if

If Varp5 <> "" and not isnull(Varp5) then
Pricelist = (Varp1 & ", " & Varp2 & ", " & Varp3 & ", " & Varp4 & ", " & Varp5)
end if

If Vard6 <> "" and not isnull(Varp6) then
Pricelist = (Varp1 & ", " & Varp2 & ", " & Varp3 & ", " & Varp4 & ", " &_
 Varp5 & ", " & Varp6)

end if

If Varp7 <> "" and not isnull(Varp7) then
Pricelist = (Varp1 & ", " & Varp2 & ", " & Varp3 & ", " & Varp4 & ", " &_
 Varp5 & ", " & Varp6 & ", " & Varp7)
end if

If Varp8 <> "" and not isnull(Varp8) then
Pricelist = (Varp1 & ", " & Varp2 & ", " & Varp3 & ", " & Varp4 & ", " &_
 Varp5 & ", " & Varp6 & ", " & Varp7 & ", "  & Varp8)
end if


If Varp9 <> "" and not isnull(Varp9) then
Pricelist = (Varp1 & ", " & Varp2 & ", " & Varp3 & ", " & Varp4 & ", " &_
 Varp5 & ", " & Varp6 & ", " & Varp7 & ", "  & Varp8 & ", " & Varp9)

end if

If Varp10 <> "" and not isnull(Varp10) then
Pricelist = (Varp1 & ", " & Varp2 & ", " & Varp3 & ", " & Varp4 & ", " &_
 Varp5 & ", " & Varp6 & ", " & Varp7 & ", "  & Varp8 & ", " & Varp9 &_
", " & Varp10)
end if


InOrder = Request("inorder")
If InOrder = "" then
InOrder=0
end if

Dim EdID, Delete
EdID=Request.Form("update")
Delete = REquest.Form("delete")
If Delete <> "" then

dim objConnD, objRSD, sqlStringD
set objConnD=Server.CreateObject("ADODB.Connection")
objConnD.ConnectionString=strconnect
objConnD.Open
Set objRSD = Server.CreateObject("ADODB.Recordset")
sqlStringD = "Select * from products where ID =" & EdID
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
  sqlString = "Select * From Products Where ID= " & EdID
  objRS.ActiveConnection = objConn
  objRS.lockType=adLockOptimistic
  objRS.Open sqlString

  objRS("Name")=Name
  objRS("Image") = Image
  objRS("Width") = Width
  objRS("Height") = Height
  objRS("Catagory") = Catagory
  objRS("Rabbits") = Rabbit
  objRS("Rodents") = Rodent
  objRS("Ferrets") = Ferret
  objRS("DescriptionLng") = Description
  objRS("Variations") = Variations
  objRS("Status") = Status
  objRS("Variation1Des") = Varlist
  objRS("Variation1Price") = Pricelist
  objRS("InOrder") = InOrder
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
  objRS2.Open "products", objConn2, , adLockOptimistic, adCmdTable


  objRS2.AddNew
  objRS2("Name") = Name
  objRS2("Image") = Image
  objRS2("Width") = Width
  objRS2("Height") = Height
  objRS2("Catagory") = Catagory
  objRS2("Rabbits") = Rabbit
  objRS2("Rodents") = Rodent
  objRS2("Ferrets") = Ferret
  objRS2("DescriptionLng") = Description
  objRS2("Variations") = Variations
    objRS2("Status") = Status
  objRS2("Variation1Des") = Varlist
  objRS2("Variation1Price") = Pricelist
  objRS2("InOrder") = InOrder
  objRS2.Update
  objRS2.Close
  set objRS2 = nothing
  objConn2.Close
  set objConn2 = nothing
  end if
end if
Response.redirect "updatep.asp"
end if

%>

