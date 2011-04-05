<% @LANGUAGE=VBSCRIPT %>
<% OPTION EXPLICIT %>
<!-- #include file="correct-domain.asp"-->
	<!-- Created: 6/23/01 8:03:58 AM -->
	<! DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2//EN">
<HTML>
<HEAD>
<SCRIPT LANGUAGE="JavaScript">
function validate() {

 if (document.apppair.email.value.length < 5) {
 alert("Please enter your e-mail address so we can contact you!");
 return false;
}
return true;
}

</script>
	<META NAME="GENERATOR" Content="ASP Express">
	<META HTTP-EQUIV="Content-Type"CONTENT="text/html;CHARSET=iso-8859-1">
	<TITLE>Washington HRS - Application to adopt a rabbit</TITLE>
<STYLE TYPE="text/css"><!--A { text-decoration: none }A:hover { text-decoration: underline }--></STYLE>
<!--#include file="google-analytics.js"-->
</head>

<BODY BGCOLOR="#ffffff" LINK="#B8860B"  ALINK="#8FBC8F" VLINK="#2E8B57">
<center><table bgcolor="#ffffff" width="100%" cellpadding="0" cellspacing="1">

<tr><td align=left>
<!--#include file="washlinks.asp"-->
</td></tr>
</table>
<table width=95% cellpadding=6 bgcolor=#ffffff border=1>
<tr><td>
<p align=center><FONT SIZE=3><b>APPLICATION TO ADOPT <br>A RABBIT FROM THE <br>WASHINGTON HOUSE RABBIT SOCIETY</b>
</FONT>
</p>

<p align=left>
Please answer the following questions and submit the form to us.  You can also e-mail your answers to 
Sandi@rabbitmeadows.org (by copy/pasting them to an e-mail message) or print it out and snail mail it to:<br>
<b>Washington HRS<br>14317 Lake City Way NE
<br>
Seattle, Washington, 98125</b>
 
<p>
The following may seem excessively long and detailed but please understand that when we place rabbits with you, we want
to be sure that you have considered as many as possible of the reasons that cause people to abandon rabbits to shelters or other homes.  Rabbits
are sensitive creatures and bond deeply to their homes and their people.  We want to do everything humanly 
possible to guarantee that when we place rabbits, they will remain in their home permanently.
<p>
<b>1. Personal Information</b><br><br><table border=1 bordercolor="burlywood" width=90%>
<form name="appwash" method="post" action="appmailwash.asp" onSubmit="return validate();">
<tr><td>
<font face="arial">Name</font> <br><input type=text name="applicantname" size="30"><br><br>
<font face="arial">Address </font><br><textarea name="address" rows="4" cols="30" wrap=soft></textarea></td>
<td>
<font face="arial">Phone</font> <br><INPUT TYPE="text" name="phone" size="15"><BR><br>
<font face="arial">Alternate Phone (if available)</font><br><INPUT TYPE="text" name="altphone" size="15"><BR><br>
<font face="arial">E-Mail (required so we can respond to you)</font><br><INPUT TYPE="text" name="email" size="20"></td></tr></table><br>

2.  Do you own or rent your home?  If you rent, do you have written permission to have a pair of rabbits in your home?
<br>
<TEXTAREA NAME="2" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>
2a. Do we have permision to visit your home?  If no please explain.
<TEXTAREA NAME="2a" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>
3. Are there any smokers in your home?<br>
<TEXTAREA NAME="3" Rows="2" cols="60"  WRAP="SOFT"></TEXTAREA><BR><br>
4. Who all is in the family (or shares your home)?  Please give the names and ages of all children, and what
part they are playing in getting rabbits (i.e., rabbits are primarily for this child, or this child wants to be able to carry
the rabbits around, etc.)<br>
<TEXTAREA NAME="4" Rows="6" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>

5. Are all of the adults in the family in agreement about getting a rabbit or rabbits?<br>
<input type="radio" name="5a" value="Yes"><b>Yes</b><br><input type="radio" name="5a" value="No"><b>No</b><br><input type="radio" name="5a" value="Notsure" ><b>Not Sure</b><br>

<textarea name="5b" rows="2" cols="60">Put additional comments (if any) here</textarea><br><br>

6.  If you have children, do you understand and agree that:
<UL>
<li>Children, even teenagers, lack the maturity to take full responsibility for the well-being of a living animal, and
<li>You will be primarily responsible for feeding, cleaning, grooming, and giving physical affection to the rabbits, for their
entire lives, even if the children lose interest in the rabbits or leave home?
</ul>
<input type="radio" name="6a" value="Yes"><b>Yes</b><br><input type="radio" name="6a" value="No"><b>No</b><br><input type="radio" name="6a" value="NotSure"><b>Not Sure</b><br>

<textarea name="6b" rows="2" cols="60">Put additional comments (if any) here</textarea><br><br>

7. Have you had a rabbit before?  Have you had any experience with rabbits?<br>
<TEXTAREA NAME="7" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>

8. Why do you want rabbits?<br>
<TEXTAREA NAME="8" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>

9. Describe what you expect rabbits in your home to be like, how you expect them to behave toward you, what
problems you think you might have, etc. (We want to be sure to educate you if your expectations are unrealistic).
<br>
<TEXTAREA NAME="9" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>

10. Do you have other animals?  If so, who/what/how old/etc. (Rabbits do fine with many animals of other species, so
 this is not usually a problem.  Some animals can be a threat to rabbits and, of course, we won't place rabbits where
such animals live.  If you do have other animals, we will work with you to teach you how to introduce them to each other).<br>
<TEXTAREA NAME="10" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>

11. Have you had other animals as an adult that you don't have now?  If so, what?  What happened to them?<br>
<TEXTAREA NAME="11" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>

11a.  Have you ever had to give up one of your animals?  If yes, please explain.<br>
<TEXTAREA NAME="11a" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>


12. Where in your home would you plan to keep your rabbits' own home ("cage")? (We recommend that it be in a family room or kitchen--somewhere 
where the people in the home tend to spend most of their time, so that even when the rabbits aren't out, they can be part 
of what's going on.)  By the way, don't buy a cage until after you've chosen your rabbits, or else get one that would house the largest of rabbits.  Also, don't
 assume that a cage you have used in the past will be considered adequate by us.<br>
<TEXTAREA NAME="12" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>


13. Do you know for certain that no one in your home is allergic to rabbits or hay (the primary item in their diet)?  Please
 have everyone in the family spend enough time with rabbits to verify this (you may come to our foster facilities for this purpose, if you wish), 
or have everyone tested by an allergist before adopting. <br>
<TEXTAREA NAME="13" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>


14. Do you have the financial ability (and stability) to cover veterinary costs that could occur at any 
time and are all of the adults in the family willing to pay for veterinary care for illnessses and for annual check-ups?  
Vet care can be quite expensive ($60.00 to $600.00 for a single episode of illness), and you need to be financially 
secure enough to handle unexpected visits.<br>
<TEXTAREA NAME="14" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>

14a. Please list the name and address of your veterinarian.
<TEXTAREA NAME="14a" Rows="2" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>



15. If you became so busy that you felt it wasn't fair to your rabbits to keep them, what would you do? (Think 
about this in the same terms as if the rabbits were children).<br>
<TEXTAREA NAME="15" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>


16. If you were to move across the country, what would you do with your rabbits?</b><br>
<TEXTAREA NAME="16" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>


17. If you were to start sharing your life with someone who was severely allergic to rabbits, or have someone
 in the family become allergic to them, how would you deal with the problem?<br>
<TEXTAREA NAME="17" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>


18. If you were to have a baby sometime in the future, would you be able to deal with keeping the rabbits in 
spite of the extra workload, and solving the problems associated with having both babies and rabbits in the house (we
 will help you with this if you need us)?<br>
<TEXTAREA NAME="18" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>


19.  Are you willing to limit any animals you get in the future to those who would not be a threat to the rabbits (no terriers, 
dachshunds, chows, pit-bulls, snakes, ferrets, etc.)?  Are you willing to limit the number of animals you have so that you could always
 afford vet care for the rabbits, should they need it?<br>
<TEXTAREA NAME="19" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>

20.Do you have time to provide the rabbits with exercise for several hours each day?
<TEXTAREA NAME="20" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>
21.Can you agree that the rabbits will not be left outside your house from dusk to dawn and when outside during the day, 
that they are constantly supervised?<br>
<TEXTAREA NAME="21" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>


22. Are there any circumstances in your life that would make it probably that you would not be able to keep
 your rabbits for their entire lives (international moves, etc.)?<br>
<TEXTAREA NAME="22" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>

23. Are you committed to keeping the rabbits, solving any problems that come up, moving only where the 
rabbits are allowed, etc., for as long as they live (usually 8 to 12 years, but possibly as much as 16 yrs.)?<br>
<TEXTAREA NAME="23" Rows="4" cols="60" WRAP="SOFT"></TEXTAREA><BR><br>





<INPUT TYPE="Submit" value="Send Application"><font face=arial size=3>&nbsp; &nbsp; &nbsp; <i>  Click button once only, then allow a few minutes for a confirmation message.</i></font>
</form>
<hr>
</td></tr>
</table></center><br>
<center><img src="links.gif" border="0"></center>
</BODY>
</HTML>
