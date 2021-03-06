<% @LANGUAGE=VBScript %>
<% Option Explicit %>
<%
    dim garbage, url, modified
    garbage = "404;http://" & Request.ServerVariables("SERVER_NAME") & ":"  & Request.ServerVariables("SERVER_PORT")
    url = Replace( Request.QueryString, garbage, "" )
    url = lcase(url) //We're on a Windows server, so file names aren't case sensitive.

    modified = Replace( url, "/rabbitrodentferret.org", "/shelter", vbTextCompare)
    modified = Replace( modified, "/washingtonhouserabbitsociety.org", "/shelter", vbTextCompare)
    modified = Replace( modified, "/barbaradeeb.org", "/memorial", vbTextCompare)
    if url = modified then
        Response.Status = "404 Not Found"
    else
        dim newlocation
        newlocation = "http://" & Request("HTTP_HOST") & modified
        Response.Status = "301 Moved Permanently"
        Response.AddHeader "Location", newlocation
        Response.Write "<html><body>The page you are looking for has moved to <a href='"
        Response.Write newlocation
        Response.Write "'>"
        Response.Write newlocation
        response.Write "</a></body></html>"
        Response.End
    end if
%>
<html>
    <head>
        <title>404 - RabbitMeadows.org</title>

        <link rel="StyleSheet" href="/shelter/style.css" type="text/css" media="screen"/>
        <!--#include file="sandiedit.js"-->
    </head>
    <body>

    <!--#include file="dropdownmenu.asp"-->
    <!--#include file="headerfile.asp"-->
    <center>
    <table cellspacing="10px">
        <tr>
            <td valign="top" align="right">
                <!--#include file="sidebar_left.asp"-->
            </td>
            <td width="540px" valign="top">
                 <div class="banner" style="text-align:center;">
                    <h1>Oops! - <code>404</code></h1>
                 </div>
                 <br />
                 <img src="/shelter/BunnyImages/snoopy2.jpg" alt="Where did it go?" style="float:left; padding:10px"/>
                 <p>
                     We can't seem to find the page you were looking for.  Maybe it moved, or you typed the address wrong, or somebody took it and hid it behind the couch.
                 </p>
                 <p>
                    Assuming it didn't get chewed up, you can probably find what you're looking for by clicking one of the links on the left or at the top of the page.
                    Or you could go back to the <a href="http://www.rabbitmeadows.org/shelter/">home page</a>.
                 </p>
                 <p>
                    Sorry about that.
                 </p>
            </td>
            <td valign="top" align="left">
                <!--#include file="sidebar_right.asp"-->
            </td>
        </tr>
        <tr>
            <td colspan="3">
                <!--#include file="footer.asp"-->        
            </td>
        </tr>
    </table>
    </center>
    </body>
</html>