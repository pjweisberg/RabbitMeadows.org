<%

Dim strSQLP, objRSP, objConnP,  orderTotal, Var

Dim strCustomerName, strUserStreet, strUserCity, strUserState, strTaxState
Dim strUserZip, strUserPhone, strUserEmail, strUserCCType, strUserCCNumber, strUserCCExpires

strSQLP="Select [CustomerName], [UserStreet], [UserCity], [UserState], " &_
"[UserZip], [UserPhone], [UserEmail], [UserCCType], [UserCCNumber], [UserCCExpires]  From Users Where UserID=" & UserID

Set objConnP=Server.CreateObject("ADODB.Connection")
objConnP.ConnectionString = Application("strConn")

objConnP.Open

set objRSP = Server.CreateObject("ADODB.Recordset")
objRSP.Open strSQLP, objConnP
If not objRSP.EOF then

strCustomerName = objRSP("CustomerName")
strUserStreet = objRSP("UserStreet")
strUserCity = objRSP("UserCity")
strUserState = objRSP("UserState")
strTaxState= Ucase(objRSP("UserState"))
strUserZip=objRSP("UserZip")
strUserPhone=objRSP("UserPhone")
strUserEmail=objRSP("UserEmail")
strUserCCType=objRSP("UserCCType")
strUserCCNumber=objRSP("UserCCNumber")
strUserCCExpires=objRSP("UserCCExpires")

strUserCCNumber="************" & Right(strUserCCNumber, 4)

End IF

objRSP.Close
Set objRSP=nothing
'-----------------------------------
If Request("UpdateQ") <> "" Then

  dim newQ, deleteProduct

  Set objRSP = Server.CreateObject("ADODB.Recordset")
  objRSP.locktype = adlockoptimistic
  objRSP.cursortype = adopenkeyset

  strSQLP= "Select OrderID, OrderQuantity From Orders " &_
  "Where UserID=" & UserID 

  objRSP.Open strSQLP, objConnP

  Do While Not objRSP.EOF  
  
  NewQ=Trim(Request("pq" & objRSP("OrderID")))
  deleteProduct = Trim(Request("pd" & objRSP("OrderID")))
   If newQ = "" OR newQ="0" Or deleteproduct <> "" then
   objRSP.Delete
   Else
    If isNumeric(NewQ) Then
    objRSP("OrderQuantity") = newQ

   End If

  End If
 objRSP.MoveNext
 Loop
 objRSP.Close
 set objRSP=Nothing
End If

%>
<html>
<head>
<title>Best Little Rabbit, Rodent and Ferret House - Register</title>
<!--#include file="google-analytics.js"-->
</head>
<body bgcolor=#ffffff LINK="#4486c4"  ALINK="#999944" vlink="#ff9533">
<center><img src="images\titlebaska.gif" width=359 height=24></center><br>
<%

strSQLP = "Select OrderID, OrderName, OrderVar, OrderQuantity, OrderPrice From Orders " &_
 "Where UserID=" & UserID & " " & "ORDER by OrderID Desc"

set objRSP = Server.CreateObject("ADODB.Recordset")
objRSP.Open strSQLP, objConnP
If objRSP.EOF Then
%>
<center><table width=80% height=100% border=0>
<tr><td align=center>

 <font face=arial size=4 color=#999933>
<b>Oops, there are no items left in your cart!  </font>
<p>
<form action="/index.asp">
<center><input type = "submit" value="Continue Shopping"></center>
</form>
<p>
</td></tr>
<tr><td align=center valigh=bottom>
<hr size=1 noshade width=100%>
<tr><td align = center valign=bottom>

<!--#include file="pgbottom.asp"-->
</td></tr>
</table></center>


<%
Else

orderTotal=0
%>
<form method="post" action="update.asp">
<input name="updateQ" type="hidden" value="1">

<input name = "UserID" type="hidden" value = "<%=UserID%>">

<center>
<table border=0 cellspacing=0 cellpadding=3 width=85% bgcolor=#bdd5eb cellspacing=0>
<tr bgcolor=#39638c><td colspan=4 align=left>
<font face=arial color=#ffffff size=4><b>Step 1:</font><font face=arial color=#ffffff> Verify your order. If you make changes, click "Update Basket."</b></font></td></tr>
<tr bgcolor=#ffffff><td colspan=4><img src="images\spacer.gif" width=1 height=10></td></tr>
<tr><td colspan=4>
  <table border=0  cellspacing=0>
   <tr bgcolor="#ffffff">
   <td colspan=4 align=center><font face=arial size=4><B>Order Summary</b></td></tr>
   <tr bgcolor=#ffffff><td align=center><font face="arial" size=3><b>Item</b></font></td>
   <td align=center><font face="arial" size=3><b>Quantity</b></FONT></TD>
    <TD ALIGN=CENTER><FONT FACE="arial" size=3><b>Each</b></FONT></TD>
    <TD ALIGN=CENTER><FONT FACE="arial" size=3><b>Total</b></FONT></TD></TR>
<%
Do While not objRSP.EOF
orderTotal = orderTotal + (objRSP("OrderPrice")) * (objRSP("OrderQuantity"))
Var = objRSP("OrderVar")
  If Var = "none" then
  Var = ""
  Else 
  Var = " - " & Var
  End If
 %>

    <tr bgcolor=#fff0c1>
   <td><font face="arial"><b>
  <%=objRSP("OrderName")%></b></font>
   &#32; &#32; 
   <font face="arial">
   <%
  Response.Write Var
  %>
   </font>
   </td>
  <td align=right>
   <input name="pq<%=objRSP("OrderID")%>" type="text" size=4 value ="<%=(objRSP("OrderQuantity"))%>">
   <input name = "pd<%=objRSP("OrderID")%>" type="checkbox" value = "1"> Delete

   </td>
  <td align=right>
  <%=formatcurrency(objRSP("OrderPrice"))%>
   </td>
   <TD align=right><%=formatcurrency((objRSP("OrderPrice")) * (objRSP("OrderQuantity")))%></td>
   </tr>   

   <%
  objRSP.MoveNext
   Loop

  objRSP.Close
  set objRSP=Nothing
  objConnP.Close
  set objConnP=Nothing

  dim Tax
  If strTaxState = "WA" then
  Tax = Round ((orderTotal * .086), 2)
  Else
  Tax = 0
  End If

   Dim Shipping
   Shipping = ShipCharges(OrderTotal)

  %>
  <tr><td colspan=4 align=right>
     <tablE>
     <tr>
      <td rowspan=4>&nbsp</td> <td align=right colspan=4><FONT FACE=ARIAL><B>Subtotal:</b></font></td><td align=right> <%=FormatCurrency(orderTotal)%></td></tr>

   <%
   If Tax <> 0 then
   %>

    <tr><td align=right colspan=4><font face=arial><b>Tax: </b></font></td><td align=right> <%=formatcurrency(Tax)%></td></tr>

   <%
   End If
   %>
    <tr><td align=right colspan=4><font face=arial><b>Shipping:*</b></font></td>
    <td align=right> <%=formatcurrency(Shipping) %></td></tr>

   <tr><td align=right colspan=4><font face=arial size=4><b>Total:</b></font></td><td align=right><b> <%=formatcurrency(orderTotal + Tax + Shipping) %></b></td></tr>
   <tr><td colspan=2>*<I>Shipping amount does not include extra charges on oversize items.  If they apply they were noted
    on the product descriptions of items you ordered and will be added later.</i></td></tr>

     </table>

  </td></tr>
    <tr><td colspan=4><table border=0 width=100% width=70%>
   <tr><td colspan=2 align=right>
        <input type = "Submit" value="Update Basket"></td>
       </form></td>
     
   </tr></table>
</td></tr>
</table>
</td></tr>
<tr bgcolor=#ffffff><td colspan=4><img src="images\spacer.gif" width=1 height=10></td></tr>

<tr bgcolor=#39638c><td colspan=4 align=left><font face=arial color=#ffffff size=4><b>
Step 2:</font><font face=arial color=#ffffff> Verify personal and payment information.  If you make changes, click "Update Personal."<br>
</b></font></td></tr>

<tr bgcolor=#ffffff><td colspan=4><img src="images\spacer.gif" width=1 height=10></td></tr>

<tr><td colspan=4>

<form method="post" action="checkout2.asp">

  <table width=100% border=0 bgcolor=#ffffff cellpadding=4 cellspacing=0>

  <tr><td>

   <font face="Arial" size="4"><b><%=strCustomerName %></b> <br>

  <font face="Courier" size="2">
  <br><b>street:</b></font>
  <font face=arial><input name="street" size=20 maxlength=50 value="<%=strUserStreet %>"></font>
  <br><font face="Courier" size="2"><b>city:</b></font>
  <font face=arial><input name="city" size=20 maxlength=50 value="<%=strUserCity %>"></font>
  <br><font face="Courier" size="2"><b>state:</b></font>
  
  <font face=arial><input name="state" size=2 maxlength=2 value="<%=strUserState %>"></font><br>
  <br><font face="Courier" size="2"><b>zip:</b></font>
  <font face=arial><input name="zip" size=20 maxlength=20 value="<%=strUserZip %>">
  </font><br>

  <font face="Courier" size="2"><b>telephone:</b></font>
  <font face=arial><input name="phone" size=12 maxlength=20 value="<%=strUserPhone %>"></font>
  <br>
  <font face="Courier" size="2">
  <b>email address:</b>  </font>
  <font face=arial><input name="email" size=30 maxlength=75 value="<%=strUserEmail%>">
  </font><br>
  <font face="Arial" size="3" >
  <p><b>Payment Information:</b>
  </font>
  <font face="Courier" size="2">
  <br><b>type of credit card:</b></font>
  <font face="arial" >  
  <input name="cctype" value="<%=strUserCCType%>"> </font> 

  <br><font face="Courier" size="2"><b>credit card number:</b></font> 
  <font face=arial><input name="ccnumber" size=20 maxlength=20 value="<%=strUserccNumber%>"></font>
  <font face="Courier" size="2">
  <br><b>credit card expires:</b> (mm/yy)</font></b>  
  <font face=arial><input name="ccexpires" size=4 maxlength=20 value="<%=strUserccExpires%>"></font>
  </td></tr>
  <tr><td align=right>
  <input type="hidden" name="UserID" value = "<%=UserID%>">
  <input type="hidden" name="updatepers" value="1">
  <input type = "submit" value="Update Personal" >

   </td></tr></table></form>

</td></tr>
<tr bgcolor=#ffffff><td colspan=4><img src="images\spacer.gif" width=1 height=10></td></tr>
<tr bgcolor=#39638c><td colspan=4 align=left><font face=arial color=#ffffff size=4><b>
Step 3:</font><font face=arial color=#ffffff> Click "Send Order" to complete your order.
</b></font></td></tr>
<tr bgcolor=#ffffff><td colspan=4><img src="images\spacer.gif" width=1 height=10></td></tr>

<tr bgcolor=#fff0c1><td colspan=4 align=right>

<form method="post" action="checkout.asp">
<input type = "hidden" name="UserID" value = "<%=UserID%>">
<input type="hidden" name="checkout" value="1">
<input type="submit" value = "Send Order">

</td></tr>
<tr bgcolor=#ffffff><td colspan=4><img src="images\spacer.gif" width=1 height=10></td></tr>
</td></tr></table>

<!--#include file="pgbottom.asp"-->

<% End If %>
</body></html>

