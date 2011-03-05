<%




Function isAlpha(n)

 If n  = "a" or n = "A" or n = "b" or n= "B" or n= "C" or n="c" then
isAlpha=true

elseif n = "d" or n = "D" or n = "e" or n= "E" or n= "f" or n="F" then
isAlpha=true

elseif n = "g" or n = "G" or n = "h" or n= "H" or n= "i" or n="I" then
isAlpha=true

elseif n = "j" or n = "J" or n = "k" or n= "K" or n= "l" or n="L" then
isAlpha=true

elseif n = "m" or n = "M" or n = "n" or n= "N" or n= "o" or n="O" then
isAlpha=true

elseif n = "p" or n = "P" or n = "q" or n= "Q" or n= "r" or n="R" then
isAlpha=true

elseif n = "s" or n = "S" or n = "t" or n= "T" or n= "u" or n="U" then
isAlpha=true

elseif n = "v" or n = "V" or n = "w" or n= "W" or n= "x" or n="X" then
isAlpha=true

elseif n = "y" or n = "Y" or n = "z"  or n="Z" then
isAlpha=true

elseif n="0" or n="1" or n="2" or n="3" or n="4" or n="5" or n="6" or n="7" or n="8" or n="9" then
isAlpha=true

else isAlpha=false
end if
end function

Function CheckAlpha(string)
dim i
For i = 1 to Len(string)
If isAlpha(MID(string, i, 1)) then
CheckAlpha = true
else 
CheckAlpha=false
exit for
end if
next
end function




Function isState(state)

If len(trim(state)) = 2 then 
isState=1
else
isState=0

End IF

End Function
 
SUB addCookie (theName, theValue)
	Response.Cookies( theName ) = theValue
	Response.Cookies( theName ).Expires = "July 31, 2001"
	Response.Cookies( theName ).Path = "/"
	Response.Cookies( theName ).Secure = FALSE
END SUB

FUNCTION fixQuotes( theString )
  fixQuotes = REPLACE( theString, "'", "''" )
END FUNCTION

FUNCTION CleanCCNum( ccnumber )
  Dim i
  FOR i = 1 TO LEN( ccnumber )
    IF isNumeric( MID( ccnumber, i, 1 ) ) THEN
      CleanCCNum = CleanCCNum & MID( ccnumber, i, 1 )
    END IF
  NEXT
END FUNCTION

FUNCTION validCCNumber( ccnumber )
  Dim isEven, digits, i, checkSum
  ccnumber = CleanCCNum( ccnumber )
  IF ccnumber = "" THEN
    validCCNumber = FALSE
  ELSE
  isEven = False
  digits = ""        
  for i = Len( ccnumber ) To 1 Step -1
  if isEven Then
    digits = digits & CINT( MID( ccnumber, i, 1) ) * 2
  Else                
    digits = digits & CINT( MID( ccnumber, i, 1) )
  End If            
  isEven = (Not isEven)
  Next
  checkSum = 0
  For i = 1 To Len( digits) Step 1
    checkSum = checkSum + CINT( MID( digits, i, 1 ) )        
  Next
  validCCNumber = ( ( checkSum Mod 10) = 0 )
  END IF
End Function



FUNCTION checkpassword
Dim objConn, sqlStringP, objRSP, UserID

  sqlStringP = "SELECT UserID, CustomerName FROM Users " &_
    "WHERE UserName='" & strUserName & "' " &_
    "AND UserPassword='" & strPassword & "'"

Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = Application("strConn")

objConn.Open
 
SET objRSP = Server.CreateObject("ADODB.Recordset")
objRSP.Open sqlStringP, objConn
  
  IF objRSP.EOF THEN 
    checkpassword = -1
  ELSE
    checkpassword = objRSP( "UserID" )
    strCustName = objRSP("CustomerName")
    UserID = checkpassword
'Response.Write UserID

    
  END IF
objRSP.Close
Set objRSP=Nothing
objConn.Close
Set objConn=Nothing
END FUNCTION

Function invalidEmail(email)
If instr(email, "@") = 0 or instr(email, ".") = 0 then
invalidEmail=True
Else
invalidEmail=False
End if
End Function

FUNCTION alreadyUser( theUsername )
Dim sqlString, objConn, objRSA
  sqlString = "SELECT UserName FROM Users " &_
    "WHERE UserName='" & fixQuotes( theUsername ) & "'"

Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = Application("strConn")

objConn.Open

  SET objRSA = Server.CreateObject("ADODB.Recordset")
	objRSA.Open sqlString, objConn
  IF objRSA.EOF THEN
    alreadyUser = FALSE
  ELSE
    alreadyUser = TRUE
  END IF
  objRSA.Close
  Set objRSA = Nothing
  objConn.Close
  Set objConn=Nothing
	
END FUNCTION

FUNCTION AddorUpdate

Dim  objConn, strSQL, objRS

  strSQL = "SELECT UserID FROM Users " &_
    "WHERE CustomerName= " & "'" & strCustName & "' " &_
    "AND UserStreet= '" & strStreet & "' " &_
    "AND UserCity= '" & strCity & "' "

Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString=Application("strConn")

objConn.Open
 
SET objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open strSQL, objConn

If objRS.EOF then

AddorUpdate="Add"
Else
UserID=objRS("UserID")
AddorUpdate="Update"
End If

objRS.Close
set objRS = Nothing
objConn.Close
Set objRS = Nothing


End Function

SUB AddUser

 dim problems, errorMSG

errorMSG = "<b>There was a problem with the information you entered.</b>  We need you to: <br>  "
problems=0

If (strNewUserName)<>"" then

  If alreadyUser(strNewUserName) then
  errorMSG = errorMSG + "Choose a different user name.  The one you picked is in use. <br>"
  problems=1
  End If

  If CheckAlpha(strNewUserName) = False then
    errorMSG=errorMSG + "Use alpha and numeric characters only for your user name<br>"
    problems=1
   End if

End IF
  

  If (strNewPassword)<>"" then

   If  CheckAlpha(strNewPassword)=False then
   errorMSG=errorMSG + "Use alpha and numeric characters only for your password<br>"
   problems =1
   End if

   If len(strNewPassword) < 6 then
   errorMSG = errorMSG + "Enter at least six alpha or numeric characters for your password<br>"
   problems=1
   end if

   If Trim(strNewUserName)=Trim(strNewPassword) then
   errorMSG=errorMSG + "Enter a password that is different from your user name <br>"
   End If

 End if

If strCustName = "" Then
  errorMSG = errorMSG + "Enter your name <br> "
  problems=1
End If

If strStreet = "" Then
  errorMSG = errorMSG + "Enter your street <br>"
  problems=1
End If

If strCity = "" Then
  errorMSG = errorMSG + "Enter your city <br>"
  problems=1
End If

If strState = "" Then
  errorMSG = errorMSG + "Enter your state <br>"
  problems=1
End If



If strZip = "" Then
  errorMSG = errorMSG + "Enter your zip code <br> "
  problems=1
End If

If stremail = "" Then
  errorMSG = errorMSG + "Enter your e-mail address <br>"
  Problems=1
End if

If stremail<>"" AND invalidEmail(stremail) = True Then
  errorMSG = errorMSG + "Enter a valid e-mail address <br>"
  Problems=1
End if

If strccNumber = "" Then
  errorMSG = errorMSG + "Enter your credit card number <br> "
  problems=1
End If

If strccNumber <> "" And Not validCCNumber(strccNumber) Then
  errorMSG = errorMSG + "Enter a valid credit card number <br> "
  problems=1
End if

If strccExp = "" Then
  errorMSG = errorMSG + "Enter your credit card expiration <br> "
  problems=1
End If

If strccExp <> "" And Not IsDate(strccExp) Then
 errorMSG = errorMSG + "Enter a valid expiration date <br> "
 problems=1
End IF



If problems= 0 then


 If strNewUserName="" then
   strNewUserName="blank"
 End If

 If strNewPassword = "" then
   strNewPassword = "blank"
 End If

 If strPhone = "" then
   strPhone = "blank"
 End If

 If strEmail = "" then
    strEmail="blank"
    End If


Call AddorUpdate
 If AddorUpdate = "Add" Then

Dim objConn, sqlString, sqlString2, objRS
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = Application("strConn")

objConn.Open

 sqlString = "Insert Into Users (" & "UserName, " & "UserPassword, " & "CustomerName, " &_
 "UserStreet, " & "UserCity, " & "UserState, " & "UserZip, " &_
 "UserPhone," & "UserEmail, " & "UserccType, " & "UserccNumber, "& "UserccExpires" & ") VALUES (" &_
  "'" & fixQuotes(strNewUserName) & "'," &_
 "'" & fixQuotes(strNewPassword) & "'," &_
 "'" & fixQuotes(strCustName) & "'," &_
 "'" & fixQuotes(strStreet) & "'," &_
  "'" & fixQuotes(strCity) & "'," &_
 "'" & fixQuotes(strState) & "'," &_
 "'" & fixQuotes(strZip) & "'," &_
 "'" & fixQuotes(strPhone) & "'," &_
 "'" & fixQuotes(strEmail) & "'," &_
 "'" & fixQuotes(strccType) & "'," &_
 "'" & fixQuotes(strccNumber) & "'," &_
 "'" & fixQuotes(strccExp) & "' " &_
 ")"

 sqlString2 = "Select @@Identity"

 objConn.Execute sqlString


Set objRS = Server.CreateObject("ADODB.Recordset")


objRS.Open sqlString2, objConn, adopenforwardonly, adlockreadonly, adcmdtext

UserID = objRS.Fields.Item(0).Value


objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn=Nothing


End if

addOrder

%>

<%  
 


ElseIF problems=1 then 


Response.Write "<font face=arial size=3 color=""red"">" & errorMSG  & "</font>"
%>
<!--#include file="regform.asp"-->
<%
Response.end
End If

End Sub



Sub addOrder

If isArray(arrCustomerCart) Then

Dim objConn, strSQL, objRS, i

Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = Application("strConn")

objConn.Open

For i=0 to UBound(arrCustomerCart, 2)
 If arrCustomerCart(orderID, i)<>"" then
   If arrCustomerCart(OrderVar, i) = "" then
   arrCustomerCart(OrderVar, i) = "none"
   End if

  Set objRS=Server.CreateObject("ADODB.Recordset")
  objRS.Locktype=adlockOptimistic
  

  strSQL = "Select UserID, OrderQuantity From Orders Where UserID =" & UserID &_
 "AND OrderName= '" & fixquotes(arrCustomerCart(Ordername, i)) &_
          "' AND OrderVar = '" & fixquotes(arrCustomerCart(OrderVar, i)) & "'"

 objRS.Open strSQL, objConn

  If objRS.EOF then
      objRS.Close

      Set objRS = Server.CreateObject("ADODB.Recordset")

  objRS.Open "Orders", objConn, , adLockOptimistic, adCmdTable

  objRS.AddNew
    objRS ("UserID") = UserID
   objRS ("OrderProductID") = arrCustomerCart(OrderID, i)
   objRS("OrderName") = arrCustomerCart(OrderName, i)
   objRS ("OrderVar") = fixquotes(arrCustomerCart(OrderVar, i))
   objRS ("OrderPrice") = arrCustomerCart(OrderPrice, i)
   objRs ("OrderQuantity") = arrCustomerCart(OrderQuantity, i)
 objRS.Update
 objRS.Close
 Set objRS = nothing

 Else

dim value
value=arrCustomerCart(OrderQuantity, i)

 
 strSQL = "UPDATE Orders Set " &_
           "OrderQuantity = " & arrCustomerCart(OrderQuantity, i) &_
           " Where OrderProductID = " & arrCustomerCart(OrderID, i)
 
objRS.Close
Set objRs=nothing

 objConn.Execute strSQL
 End If
end if
 
 Next

objConn.Close
Set objConn=Nothing
End If

End Sub


Function ShipCharges(theTotal)

If theTotal  > 0 AND theTotal <=  30.00 then
ShipCharges = 5.95
End If
If theTotal  > 30.00 AND theTotal <= 60.00 then
ShipCharges= 7.95
End If
If theTotal > 60.0 AND theTotal <= 90.00 then
ShipCharges = 9.95
End If
If theTotal > 90.00 then
ShipCharges =11.95

End If

End Function

%>

