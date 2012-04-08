<script type="text/javascript">

    /***********************************************
    * AnyLink Drop Down Menu-  Dynamic Drive (www.dynamicdrive.com)
    * This notice MUST stay intact for legal use
    * Visit http://www.dynamicdrive.com/ for full source code
    ***********************************************/

    //Contents for menu 1

    var menurab = new Array()
    menurab[0] = '<a href="/shelter/Adoptedcurrent.asp" onmouseover="changeContentPure(td1,link1rab)"  onmouseout="changeContentPure(td1,msg100)" >Adoptions</a>'

    menurab[1] = '<a href="/shelter/vets.asp?animal=1" onmouseover="changeContentPure(td1,link2rab)"  onmouseout="changeContentPure(td1,msg100)" >Vet Referral</a>'
    menurab[2] = '<a href="/shelter/sites.asp" onmouseover="changeContentPure(td1,link3rab)"  onmouseout="changeContentPure(td1,msg100)">Links</a>'
    menurab[3] = '<a href="/shelter/volunteerrab.asp" onmouseover="changeContentPure(td1,link4rab)"  onmouseout="changeContentPure(td1,msg100)">Volunteer</a>'
    menurab[4] = '<a href="/shelter/donate.asp?name=rabbits" onmouseover="changeContentPure(td1,link5rab)"  onmouseout="changeContentPure(td1,msg100)">Donate</a>'

    var menumea = new Array()
    menumea[0] = '<a href="/shelter/RabbitMeadowsSanctuary.asp" onmouseover="changeContentPure(td1,link1mea)" onmouseout="changeContentPure(td1,msg100)" >Our Sanctuary</a>'
    menumea[1] = '<a href="/shelter/volunteer.asp?name=Sanctuary Rabbits" onmouseover="changeContentPure(td1,link2mea)" onmouseout="changeContentPure(td1,msg100)" >Volunteer</a>'
    menumea[2] = '<a href="/shelter/donate.asp?name=sanctuary rabbits" onmouseover="changeContentPure(td1,link3mea)" onmouseout="changeContentPure(td1,msg100)" >Donate</a>'

    var menurod = new Array()
    menurod[0] = '<a href="/shelter/RodentCurrent.asp" onmouseover="changeContentPure(td1,link1rod)"  onmouseout="changeContentPure(td1,msg100)" >Adoptions</a>'
    menurod[1] = '<a href="/shelter/vets.asp?animal=2" onmouseover="changeContentPure(td1,link2rod)"  onmouseout="changeContentPure(td1,msg100)" >Vet Referral</a>'
    //menurod[2] = '<a href="/shelter/comingsoon.asp?name=Rodent Links" onmouseover="changeContentPure(td1,link3rod)"  onmouseout="changeContentPure(td1,msg100)" >Links</a>'
    menurod[2] = '<a href="/shelter/volunteer.asp?name=Rodents" onmouseover="changeContentPure(td1,link4rod)"  onmouseout="changeContentPure(td1,msg100)" >Volunteer</a>'
    menurod[3] = '<a href="/shelter/donate.asp?name=rodents" onmouseover="changeContentPure(td1,link5rod)"  onmouseout="changeContentPure(td1,msg100)" >Donate</a>'

    var menugui = new Array()
    menugui[0] = '<a href="/shelter/GuineaCurrent.asp" onmouseover="changeContentPure(td1,link1gui)"  onmouseout="changeContentPure(td1,msg100)" >Adoptions</a>'
    menugui[1] = '<a href="/shelter/vets.asp?animal=4" onmouseover="changeContentPure(td1,link2gui)" onmouseout="changeContentPure(td1,msg100)">Vet Referral</a>'
    //menugui[2] = '<a href="/shelter/comingsoon.asp?name=Guinea Pig" onmouseover="changeContentPure(td1,link3gui)"  onmouseout="changeContentPure(td1,msg100)">Links</a>'
    menugui[2] = '<a href="/shelter/volunteer.asp?name=Guinea Pigs" onmouseover="changeContentPure(td1,link4gui)"  onmouseout="changeContentPure(td1,msg100)">Volunteer</a>'
    menugui[3] = '<a href="/shelter/donate.asp?name=guinea pigs" onmouseover="changeContentPure(td1,link5gui)"  onmouseout="changeContentPure(td1,msg100)">Donate</a>'

    var menudon = new Array()
    menudon[0] = '<a href="/shelter/donate.asp" onmouseover="changeContentPure(td1,link1don)"  onmouseout="changeContentPure(td1,msg100)" >Donate</a>'

    var menumem = new Array()
    menumem[0] = '<a href="http://barbaradeeb.org/memorial/" onmouseover="changeContentPure(td1,link1mem)"  onmouseout="changeContentPure(td1,msg100)">Memorial</a>'
    menumem[1] = '<a href= "http://barbaradeeb.org/memorial/philosophy.aspx" onmouseover="changeContentPure(td1,link2mem)"  onmouseout="changeContentPure(td1,msg100)">Life and Work</a>'
    menumem[2] = '<a href="http://barbaradeeb.org/memorial/publications.aspx" onmouseover="changeContentPure(td1,link3mem)"  onmouseout="changeContentPure(td1,msg100)">Publications</a>'

</script>

<!--beginning of row 1, rabbit section (all section cells are a main 2 cells inside, a title cell and main info cell containing a table) -->
<table bgcolor="#FFFFFF" width="130" height="65" cellpadding="0" cellspacing="0" border="0">
	<tr><td valign="top">
		<table class="ltgreen" cellpadding=4 cellspacing=2 width="130">
		<tr><td class="dkgreen" align="center" height="15"><a class="dkgreen" href="/shelter/donate.asp"  onMouseover="changeContentPure('td1',msg2don)" onMouseout="changeContentPure('td1',msg100);">Donate</a></td></tr>

		<tr><td id="td10" bgcolor="#FFFFFF" height="40">

				<img src="/shelter/pics/donatelink.gif" align="left" border="0" width="41" height="46" vspace="2"> <script type="text/javascript">
				                                                                                              document.write(msg1don) </script>

		</td></tr>
		</table>
	</td></tr>
</table><br>

<table bgcolor="#FFFFFF" width="130" height="65" cellpadding="0" cellspacing="1" border="0">
	<tr><td valign="top">
		<table class="ltorange" cellpadding=4 cellspacing=2 width="130">
		<tr><td class="dkorange" align="center" height="15"><a class="dkorange" href="/shelter/AdoptedCurrent.asp" onMouseover="dropdownmenu(this, event, menurab, '90px'); changeContentPure('td1',msg2rab)" onMouseout="changeContentPure('td1',msg100);delayhidemenu()">Rabbits</a></td></tr>

		<tr><td id="td2" bgcolor="#FFFFFF" height="40">
				<img src="/shelter/pics/rabadoptlink.jpg" align="left" border="0" vspace="2"> <script type="text/javascript">
				                                                                         document.write(msg1rab)</script>

		</td></tr>
		</table>
	</td></tr>
</table>

<br>


<table bgcolor="#FFFFFF" width="130" height="65" cellpadding="0" cellspacing="0" border="0">
	<tr><td valign="top">
		<table class="ltblue" cellpadding=4 cellspacing=2 width="130">
		<tr><td class="dkblue" align="center" height="15"><a class="dkblue" href="/shelter/RodentCurrent.asp" onMouseover="dropdownmenu(this, event, menurod,'90'); changeContentPure('td1',msg2rod)" onMouseout="changeContentPure('td1',msg100);delayhidemenu()">Rodents</a></td></tr>

		<tr><td id="td7" bgcolor="#FFFFFF" height="40">

				<img src="/shelter/pics/rodentlink.jpg"  align="left" border="0" width="42" height="41" vspace="2"> <script type="text/javascript">
				                                                                                               document.write(msg1rod) </script>

		</td></tr>
		</table>
	</td></tr>
</table><br>

<table bgcolor="#FFFFFF" width="130" height="65" cellpadding="0" cellspacing="0" border="0">
	<tr><td valign="top">
		<table class="ltgreen" cellpadding=4 cellspacing=2 width="130">
		<tr><td class="dkgreen" align="center" height="15"><a class="dkgreen" href="/shelter/GuineaCurrent.asp" onMouseover="dropdownmenu(this, event, menugui,'90px'); changeContentPure('td1',msg2gui)" onMouseout="changeContentPure('td1',msg100);delayhidemenu()">Guinea Pigs</a></td></tr>

		<tr><td id="td9" bgcolor="#FFFFFF" height="40">

				<img src="/shelter/pics/guinealink.jpg" align="left" border="0" width="45" height="42" vspace="2"> <script type="text/javascript">
				                                                                                              document.write(msg1gui) </script>

		</td></tr>
		</table>
	</td></tr>
</table><br/>


<table bgcolor="#FFFFFF" width="130" height="65" cellpadding="0" cellspacing="1" border="0">
	<tr><td valign="top">
		<table class="ltpurple" cellpadding=4 cellspacing=2 width="130">
		<tr><td class="dkpurple" align="center" height="15"><a class="dkpurple" href="/shelter/RabbitMeadowsSanctuary.asp" onMouseover="dropdownmenu(this, event, menumea,'90px'); changeContentPure('td1',msg2mea)" onMouseout="changeContentPure('td1',msg100);delayhidemenu()">Rabbit Meadows</a></td></tr>

		<tr><td id="td4" bgcolor="#FFFFFF" height="40">
								<img src="/shelter/pics/rmslink.jpg" align="left" border="0" vspace="2"> <script type="text/javascript">
								                                                                    document.write(msg1mea) </script>

		</td></tr>
		</table>
</td></tr>
</table><br/>

<table bgcolor="#FFFFFF" width="130" height="65" cellpadding="0" cellspacing="0" border="0">
	<tr><td valign="top">
		<table class="ltorange" cellpadding=4 cellspacing=2 width="130">
		<tr><td class="dkorange" align="center" height="15"><a class="dkorange" href="http://barbaradeeb.org/memorial/" onMouseover="dropdownmenu(this, event, menumem,'90px'); changeContentPure('td1',msg2mem)" onMouseout="changeContentPure('td1',msg100);delayhidemenu()">Memorial</a></td></tr>

		<tr><td id="td8" bgcolor="#FFFFFF" height="40">

				<img src="pics/Barbdeeblink.jpg" align="left" border="0" width="43" height="42" vspace="2"> <script type="text/javascript">
				                                                                                                document.write(msg1mem) </script>

		</td></tr>
		</table>
	</td></tr>
</table>
