<% @LANGUAGE=VBScript %>
<% Option Explicit %>
<%
    Response.Status = "301 Moved Permanently"

    dim address
    address=request.servervariables("HTTP_HOST")
    address=lcase(address)
    If address="www.barbaradeeb.org" or address="barbaradeeb.org" then
        Response.AddHeader "Location", "http://BarbaraDeeb.org/BarbaraDeeb.org/"
    Elseif address="www.rabbitmeadowssanctuary.org" or address="rabbitmeadowssanctuary.org" then
        Response.AddHeader "Location", "http://www.rabbitmeadows.org/shelter/RabbitMeadowsSanctuary.asp"
    Else
        Response.AddHeader "Location", "http://www.rabbitmeadows.org/shelter/"
    End if
%>
<html>
    <head>
        <title>Rabbit Meadows</title>
    </head>
    <body>
        Please go to our <a href="http://www.rabbitmeadows.org/shelter/">home page</a> if you are not redirected automatically.
    </body>
</html>