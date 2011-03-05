<% @LANGUAGE=VBSCRIPT %>
<% OPTION EXPLICIT %>
<%
Dim strquestion
Dim question12, question13, question14, question15, question16, question17, question18, question19, question20, question21
Dim question2, question2a, question3, question4, question5, question6, question7, question8, question9, question10, question11, answer5a, answer6a
Dim question11a, question14a, question22, question23, question24

question2="2. Do you own or rent your home?  If you rent, do you have permission to have a pair of rabbits in your home?<br>"
question2a="2a. Do we have permission to visit your home?"
question3="3. Are there any smokers in your home?<br>"
question4="4. Who all is in the family(or shares your home?)<BR>"
question5= "5. Are all of the adults in the family in agreement about getting a pair of house rabbits?<br>"
If Request("5a")="Yes" then 
answer5a = "<b>Yes.  </b>"
elseif Request("5a")="No" then
answer5a="<b>No.  </b>"
elseif REquest("5a")="NotSure" then
answer5a="<B>Not Sure. </B>"
else
answer5a="<b>Not answered.  </b>"
end if

question6="6. If you have children, do you agree that you will be primarily responsible for the care of the rabbits?<br>"

If Request("6a")="Yes" then 
answer6a = "<b>Yes.  </b>"
elseif Request("6a")="No" then
answer6a="<b>No.  </b>"
elseif REquest("6a")="NotSure" then
answer6a="<B>Not Sure. </B>"
else
answer6a="<b>Not answered.  </b>"
end if

question7="7. Have you had a rabbit before?  Have you had any experience with rabbits?<br>"
question8="8.  Why do you want rabbits?<br>"
question9="9. Describe what you expect rabbits in your home to be like.<br>"
question10= "10. Do you have other animals?  Describe them.<br>"
question11="11. Have you had other animals as an adult that you don't have now?  What happened to them?<br>"
question11a="11a. have you ever had to give up an animal? Please explain<br>"
question12="12. Where in your home would you plan to keep your rabbit's own home?<br>"
question13 = "13. Do you know for certain that no one in your home is allergic to rabbits or hay?<br>"
question14 = "14. Do you have the financial ability to cover veterinary costs?<BR>"
question14a="14a. List the name and address of your veterinarian.<br>"
question15="15. If you became so busy that you felt it wasn't fair to your rabbits to keep them, what would you do?<br>"
question16="16. If you were to move across the country, what would you do with your rabbits?<BR>"
question17="17. What would you do if you were to start sharing your life with someone who was severely allergic to rabbits?<BR>"
question18="18.  If you were to have a baby sometime in the future, would you be able to handle the extra workload?<BR>"
question19="19. Are you willing to limit any future animals to those who would not be a threat to the rabbits?<BR>"
question20= "20. Do you have time to provide exercise?<br>"
question21= "21. Will rabbits outside be supervised?<br>"
question22= "22. Are there any circumstances that would make it probable that you won't be able to keep the rabbits for their entire lives?<BR>"
question23 = "23.  Are you committed to keeping the rabbits for as long as they live?<BR>"


strquestion=Request("applicantname") & "<br>" & Request("address") & "<br>"
strquestion=strquestion & Request("phone") & "<br>" & Request("altphone") & "<br>" & Request ("email") & "<p>"
strquestion=strquestion & question2 & "<B>" & Request("2") & "</b><p>" & question3 & "<B>" & Request("3")  & "</b><p>"
strquestion=strquestion & question2a & "<B>" & Request("2a") & "</b><p>"
strquestion=strquestion & question4 & "<b>" & Request("4") & "</b><p>" & question5 & answer5a & "<B>" & Request("5b")  & "</b><p>"
strquestion=strquestion & question6 & answer6a & "<B>" & Request("6b") & "</b><p>" & question7 & "<b>" & REquest("7")  & "</b><p>"
strquestion=strquestion & question8 & "<B>" & Request("8") & "</B><p>" & question9 & "<B>" & Request("9") & "</b><p>"
strquestion=strquestion & question10 & "<B>" & Request("10") & "</B><p>" & question11 & "<B>" & Request("11") & "</b><p>"
strquestion=strquestion & question11a & "<B>" & Request("11a") & "</b><P>" 
strquestion=strquestion & question12 & "<B>" & Request("12") & "</B><p>" & question13 & "<B>" & Request("13") & "</b><p>"
strquestion=strquestion & question14 & "<B>" & Request("14") & "</B><p>" 
strquestion=strquestion & question14a & "<B>" & Request("14a") & "</b><P>" & question15 & "<B>" & Request("15") & "</b><p>"
strquestion=strquestion & question16 & "<B>" & Request("16") & "</B><p>" & question17 & "<B>" & Request("17") & "</b><p>"
strquestion=strquestion & question18 & "<B>" & Request("18") & "</B><p>" & question19 & "<B>" & Request("19") & "</b><p>"
strquestion=strquestion & question20 & "<B>" & Request("20") & "</B><p>" & question21 & "<B>" & Request("21") & "</b><p>"
strquestion=strquestion & question22 & "<B>" & Request("22") & "</B><p>" & question23 & "<B>" & Request("23") & "</b>"




	
Dim objCDOSYSMail, objCDOSYSCon

Set objCDOSYSMail  = Server.CreateObject ("CDO.Message") 
Set objCDOSYSCon = Server.CreateObject("CDO.Configuration")
'Out going SMTP server
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")="smtp.winisp.net"
objCDOSYSCon.Fields("http//schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
objCDOSYSCon.Fields("http//schemas.microsoft.com/cdo/configuration/sendusing") = 2 
  objCDOSYSCon.Fields("http//schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
    
objCDOSYSCon.Fields.Update

'Update the CDOSYS Configuration
Set objCDOSYSMail.Configuration= objCDOSYSCon

objCDOSYSMail.From="Rebeccad@winisp.net"
objCDOSYSMail.To="houserabbit@charter.net"
objCDOSYSMail.Subject="Adoption Apps"
objCDOSYSMail.HTMLBody="<html><body>"  & strquestion & "</body></html>"
objCDOSYSMail.ReplyTo=Request("email")
objCDOSYSMail.Send

set objCDOSYSMail=Nothing
set objCDOSYSCon=Nothing

'objMailer.FromName="Adoption Apps"
'objMailer.FromAddress="Rebeccad@winisp.net"

'objMailer.Subject="New app to adopt a rabbit"
'objMailer.ContentType="text/html"
'objMailer.BodyText="<html><body>"  & strquestion & "</body></html>"

'objMailer.RemoteHost="mail.winisp.net"
'objMailer.AddRecipient "Rebecca", "Rebecca@winisp.net"

'objMailer.ReplyTo=Request("email")

%>
<html>
<body BGCOLOR="#ffffff" LINK="#B8860B"  ALINK="#8FBC8F" VLINK="#2E8B57"><font face=arial>
<br>
<!--#include file="washlinks.asp"-->

<hr>
<br>
<font face=arial size=3>Thank you!<br>
Your application has been sent to the Washington HRS and you will be contacted soon.  

<p>If you don't hear from HRS within a day or two, please send an email asking about the status of your application to: 
<a href="mailto: Sandi@rabbitrodentferret.org">Sandi@rabbitrodentferret.org</a> as 
on rare occasions there are problems with the with the program that sends completed applications. </font>

</body>
</html>







