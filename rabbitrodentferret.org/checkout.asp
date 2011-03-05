<% @Language=VBScript %>
<% Option Explicit %>
<%  REsponse.Buffer="true" %>
<!--#include file="adovbs.inc" -->

<%

Dim objConn, strSQL, UserID, objRS, NewOrderID, sqlString2
UserID = Request("UserID")

Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Connectionstring=Application("strConn")
Set ObjRS=CreateObject("ADODB.Connection")

objConn.Open

 strSQL = "Insert Into MasterOrders (" & "OrderUserID, " & "OrderStatus, " &_

 "OrderEntryDate" & ") VALUES (" &_
  UserID &_
  ", 0, " &_
 "'" & Now() &_ 
 "')"
Response.Write strSQL

 sqlString2 = "Select @@Identity"

 objConn.Execute strSQL


Set objRS = Server.CreateObject("ADODB.Recordset")


objRS.Open sqlString2, objConn, adopenforwardonly, adlockreadonly, adcmdtext

NewOrderID = objRS.Fields.Item(0).Value



objRS.Close
Set objRS = Nothing
'objConn.Close
'Set objConn=Nothing
'objConn.Open

objConn.BeginTrans





strSQL = "Insert into process (" &_
"OrdID, " &_
"OrdProductID, " &_
"OrdName, " &_
"OrdVar, " &_
"OrdQuantity, " &_
"OrdPrice, " &_
"OrdUserID, " &_
"OrdEntryDate, " &_
"OrdStatus, " &_
"OrdNumber, " &_
"OrdPrint " &_
") select " &_

"OrderID, " &_
"OrderProductID, " &_
"OrderName, " &_
"OrderVar, " &_
"OrderQuantity, " &_
"OrderPrice, " &_
"UserID, " &_
"Now(), " &_
"0, " &_
 NewOrderID &_
", True" &_
 

" From orders Where " &_
"UserID = " & UserID

'Response.Write strSQL

objConn.Execute strSQL

'Empty cart

strSQL = "Delete from orders " &_
"Where UserID=" & UserID

objConn.Execute strSQL

objConn.CommitTrans

objConn.close
set objConn=nothing

Dim objMailer

Set objMailer = Server.CreateObject ("SMTPsvg.Mailer") 
objMailer.FromName="Orders"
objMailer.FromAddress="Rebecca@RabbitMeadows.org"

objMailer.Subject="New Order"
objMailer.BodyText="A new order has been placed at " & Now & ". " & " The order number is " & NewOrderID & "."
objMailer.RemoteHost="mail.rabbitrodentferret.org"
objMailer.AddRecipient "Sandi", "Rebecca@RabbitMeadows.org"
objMailer.AddRecipient "Rebecca", "hrabbit@ix.netcom.com"
Set objMailer=Nothing


Set objMailer = Server.CreateObject ("SMTPsvg.Mailer") 
objMailer.FromName="Orders"
objMailer.FromAddress="hrabbit@ix.netcom.com"

objMailer.Subject="New Order"
objMailer.BodyText="A new order has been placed. " 
objMailer.RemoteHost="mail.earthlink.net"
objMailer.AddRecipient "Sandi", "Rebecca@RabbitMeadows.org"
objMailer.AddRecipient "Rebecca", "hrabbit@ix.netcom.com"
Set objMailer=Nothing

session("cart")=""

Response.Redirect "thankyou.asp"
%>
