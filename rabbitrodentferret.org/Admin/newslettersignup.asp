<% @LANGUAGE=VBScript %>
<% Option Explicit %>
<% Server.ScriptTimeout=150%>

<SCRIPT language="JavaScript">
<!--hide
function thankyouwindow(page)
{
window.open(page,'indexwin','scrollbars=yes,toolbar=yes,menubar=yes,location=yes,status=yes,resizable=yes');
}
//-->
</SCRIPT>

<%
Dim strquestion

strquestion=Request("applicantname") & "<br>" & Request("email") & "<br>"

Dim MyMail
Set MyMail = Server.CreateObject("Persits.MailSender") 
MyMail.Host = "sendmail.brinkster.com" 
MyMail.body = "<html><body>"  & strquestion & "</body>" 
MyMail.IsHTML = True 
MyMail.From = "webmaster@rabbitrodentferret.org" 
MyMail.Username = "webmaster@rabbitrodentferret.org" 
MyMail.Password = "climber" 
MyMail.AddAddress "houserabbit@clearwire.net"
MyMail.Subject = "Newsletter Sign Up" 
MyMail.AddCC "houserabbit@clearwire.net" 



if MyMail.send then 
%>

javascript:articlewindow('article_index.asp')

<%
else 
Response.Write "<font face=arial size=3> Thank you!<br>Unfortunately, for some reason there was an error when your application was sent."
Response.Write "  Technical support has been notified of the problem.  Please use your browser's back button to return to "
Response.Write "the completed form and either print and send it or e-mail it directly to co-hrs@comcast.net.  We apologize for the inconvenience.</td></tr>" & err.description
end if
set MyMail = Nothing 
%>
</body>




<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>