<% @LANGUAGE=VBScript %>
<% Option Explicit %>
<%
    dim newlocation
    newlocation = "http://www.rabbitmeadows.org/shelter/vets.asp?animal=1"
    Response.Status = "301 Moved Permanently"
    Response.AddHeader "Location", newlocation
    Response.Write "<html><body>The page you are looking for has moved to <a href='"
    Response.Write newlocation
    Response.Write "'>"
    Response.Write newlocation
    response.Write "</a></body></html>"
    Response.End
%>