<% @LANGUAGE=VBSCRIPT %>
<% OPTION EXPLICIT %>
<!--#Include file = "adovbs.inc"-->
<%

Function getStatus(order)
  if order=3 then
  getStatus="shipped"
  elseif order=0 then
  getStatus="pending"
  elseif order=2 then
  getStatus="out-of-stock"
  Else 
  getStatus=""
  end if
end Function

Function TaxAmount(stateRes, amount)
	If stateRes="WA" then
	TaxAmount=formatNumber(amount * .082)
	else TaxAmount=0.00
	end if
End Function

Dim OrderNumber, OrdEntryDate, CustomerName
Dim UserStreet, UserCity, UserState, UserZip, Date
Dim UserPhone, UserCCType, UserCCNumber, Shipping

OrderNumber=Request("OrderNumber")
OrdEntryDate=Request("OrdEntryDate")
CustomerName=Request("CustomerName")
UserStreet=Request("UserStreet")
UserCity=Request("UserCity")
UserState=Request("UserState")
UserZip=Request("UserZip")
Date=Request("Date")
UserPhone=Request("UserPhone")
UserCCType=Request("UserCCType")
UserCCNumber=Request("UserCCNumber")
Shipping=Request("Shipping")
If not isnumeric(Shipping) then
Shipping=formatCurrency(0)
else
Shipping=formatCurrency(Shipping)
end if

Dim objConn, strSQL, objRS
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = Application("strConn")

If OrderNumber <> "" then
objConn.Open

strSQL="Select * From Process Where OrdNumber = " & OrderNumber & " Order by OrdID"
Set objRS= Server.CreateObject("ADODB.Recordset")

objRS.cursortype=adopenstatic
objRS.activeconnection=objConn
objRS.Open strSQL
%>
<html>
<body>
<table border=0 cellpadding=5 cellspacing=2 width=90%>
<tr><td colspan=2><font size=4 face=arial><b>Best Little Rabbit, Rodent & Ferret House</b></font>
<font size=2 face=arial><br>14325 Lake City Way NE<br>Seattle, WA 98125<br>(206) 365-9105<br>
http://www.RabbitRodentFerret.com<br>
<hr noshade size=1></font></td></tr>

<tr><td colspan=2 align=center><font face=arial size=3><b>Order Number: </b></font>
<font face=arial><%=OrderNumber%><br>
<font face=arial size=3><b>Order Date: </b></font>
<font face=arial><%=Date%></td></tr>
<tr><td bgcolor="#D8D8D8">
<font face=arial size=3><b>Shipping Address: </b></font></td>
<td bgcolor="#D8D8D8"><font face=arial><b>Payment Method:</b></font></td></tr>
<tr><td><font face=arial>
<%=CustomerName%><br>
<%=UserStreet%><br>
<%=UserCity%>, &nbsp;
<%=UserState%><br>
<%=UserZip%><br>
<%=UserPhone%></font></td>
<td><font face=arial><b>Type: </b>
<%=UserCCType%><br><b>Last Five: </b>
<%=UserCCNumber%></font></td></tr>

</table>
<table width=90% cellpadding=5 cellspacing=2>
<tr><td bgcolor="#D8D8D8" colspan=6 align=center>
<font face=arial size=3><b>Items Ordered </b></font></td></tr>
<tr><td><b>Order ID</b></td><td><b>Item</b></td><td><b>Quantity</b></td>
<td align=right><b>Price</b></td><td align=right><b>Total</b></td><td><b>Status</b></td></tr>

<%
dim status, subTotal, finalTotal, tax
subTotal = 0
finalTotal=0
Do While not objRS.EOF


status=getStatus(objRS("OrdStatus"))

If status <> "pending" and objRS("OrdPrint")=true  then
%>
<tr><td><%=objRS("OrdID")%></td>
<td><%=objRS("OrdName")%>&nbsp;-&nbsp;<%=objRS("OrdVar")%></td>
<td align=center><%=objRS("OrdQuantity")%></td>
<td align=right><%=objRS("OrdPRice")%></td>
<td align=right><%=objRS("OrdQuantity")*objRS("OrdPrice")%></td>
<td><%=status%></td></tr>
<%
If objRS("OrdStatus")=3 then
subTotal=subTotal + (objRS("OrdQuantity") * objRS("OrdPrice"))
end if
end if
%>

<%
objRS.Movenext
Loop
objRS.close
set objRS=nothing
objConn.close
set objConn=nothing
%>
<tr><td colspan=4 align=right><b>Sub-Total</b></td>
<td align=right><%=formatNumber(subTotal)%></td></tr>
<%
tax=TaxAmount(UserState, subTotal)
If tax > 0 then
%>
<tr><td colspan=4 align=right><b>Tax</b></td>
<td align=right><%=tax%></td></tr>

<%
end if

%>
<tr><td colspan=4 align=right><b>Shipping</b></td>
<td align=right><u><%=formatNumber(Shipping)%></u></td></tr>
<tr><td colspan=4 align=right><b>Total</b></td>
<td align=right><b><%=FormatCurrency(SubTotal+tax+Shipping)%></b></td></tr>
</table>
</body>
</html>
<%
else
REsponse.Write "There is nothing to print.  You must access this page from the orders page."
End if
%>
