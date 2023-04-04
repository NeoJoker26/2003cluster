<%@ Language=VBScript %> 
<% Option Explicit %>
<% dim scope, from, playerid
scope = Request.QueryString("scp")
if scope = "" then scope = "1,2,3,4,5,6,7"
from = Request.QueryString("from")
playerid = Request.QueryString("pid")

%>

<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="Author" content="Trevor Scallan">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<title>GoS-DB Players</title>
<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
div {margin:0; padding:0; border:none; text-align:left; font-size:11px; vertical-align:top;}
#container {width:980px;}
#left {float:left; width:251px;}
#right {float:right; width:729px;}
#right-left {float:left; width:515px; margin:12px;}
#right-right {float:right; width:190px;}
#bottom {clear:both; width:100%px; margin:12px 0;}

#contrib_head {border:1px solid #c0c0c0; padding:4px 6px 4px; margin-top:12px; background: #fdfdfd; width:500px}

#ajaxplayerdetails p {margin: 0 0 3px 0; padding: 0;  font-size: 11px; font-weight:normal; text-align: left; }
#ajaxplayerdetails .name {float:left; clear:both; line-height: 20px; margin: 0 0 9 0; padding: 2 4 2 4; color: #ffffff; background-color: #404040; font-size: 14px; font-weight:bold;}
#ajaxplayerdetails .dob {float:right; clear:both; line-height: 20px; margin: 0 0 9 0; padding: 2 4 2 4; color: #ffffff; background-color: #404040; font-size: 12px; font-weight:bold;}

.matchlist td {text-align:left; margin: 0; padding: 0 4 0 4;  font-family: "Trebuchet MS",helvetica,verdana,arial,sans-serif; font-size: 11px; }

#centerpanel_manager {display: none; border: 1px solid #202020; background-color: #e3eee3; padding: 6px 12px}

-->
</style>

<script type="text/javascript"  src="jquery/jquery-1.11.1.min.js"></script>
<script>
$(document).ready(function(){

	$('#ajaxplayerdetails').html('<p style="margin-top:12;">Calculating PAFC record ... <img border="0" src="images/ajax-loader.gif">'); 
	var parameters = {playerid : '<%response.write(playerid)%>', scp : '<%response.write(scope)%>'};
	$.ajax({
  		url: "gosdb-getplayerdetails-full.asp",
  		data: parameters,
  		cache: false,
  		success: function(html){
  			var textsplit = html.split("^");
   			$("#ajaxplayerlist").html(textsplit[0]);
   			$("#ajaxplayerdetails").html(textsplit[1]);
  			}
	});

    $('#ajaxplayerdetails').on('click','.season',function() {
        $(this).append('<img style="position:absolute; left:70px; border:0;" src="images/ajax-loader.gif">');
    });
    
    $('#ajaxplayerlist').on('click','.manager_name', function(){
    	var temp1 = $(this).attr('id');
		var temp2 = temp1.substring(4)
		var managerids = temp2.split("-") 
		var ajaxparm = "id1=" + managerids[0] + "&id2=" + managerids[1] + "&source=player"
		$('#centerpanel_manager').load('gosdb-getmanagertext.asp?' + ajaxparm);
		$("#centerpanel_player").hide('slow');
		$("#centerpanel_manager").show('slow');
	});
	
	$("#centerpanel_manager").on("click",".close", function(){
	$("#centerpanel_manager").hide('slow');
	$("#centerpanel_player").show('slow');

	});
		   
	$("img").on("contextmenu",function(){
       return false;
    });
    
});
</script>

<%
' *** Ajax Overview ***
' The search player name box calls GetPlayerlist which fires gosdb-getplayerspageplayerlist.asp
'  * gosdb-getplayerspageplayerlist.asp produces list of players with link to GetPlayer, which fires gosdb-getplayerdetails-full.asp
'  * gosdb-getplayerdetails-full.asp produces season list with [+] using Toggle2 and id's based on 'tag', which fires gosdb-getplayerspagematchlist.asp
'  * gosdb-getplayerspagematchlist.asp produces a match list with [+] using Toggle3 and id's based on 'mattag', which fires gosdb-getmatchdetails1.asp
%>

<script language="javascript">

function roll_over(img_name, img_src)
   {
   document[img_name].src = img_src;
}

function Toggle2(item,scope) {
   obj=document.getElementById(item);
   objtr=document.getElementById("tr" + item);
   visible=(obj.style.display!="none")
   key=document.getElementById("x" + item);

   if (visible) {
     obj.style.display="none";
     objtr.style.backgroundColor="#ffffff";
     key.innerHTML='[+]';

   } else {

   	  objplayer=document.getElementById("d1" + item);
   	  objyears=document.getElementById("d2" + item);

      GetDetails2(objplayer.innerHTML,objyears.innerHTML,scope);
 
      obj.style.display="block";
      objtr.style.backgroundColor="#e0f0e0";
      key.innerHTML='[-]';
   }
}

var xmlHttp

function GetDetails2(str1,str2,str3)
{ 
xmlHttp=GetXmlHttpObject();

if (xmlHttp==null)
  {
  alert ("Sorry, your browser does not support this function.");
  return;
  }
document.body.style.cursor='wait';        
var url="gosdb-getplayerspagematchlist.asp";
url=url+"?p="+str1;
url=url+"&y="+str2;
url=url+"&scp="+str3;
url=url+"&sid="+Math.random();

xmlHttp.onreadystatechange=stateChanged;
xmlHttp.open("GET",url,true);
xmlHttp.send(null);
document.body.style.cursor='auto';   
}

function stateChanged() 
{ 
if (xmlHttp.readyState==4)
   { 
   obj.innerHTML = xmlHttp.responseText;
   }
}

function GetXmlHttpObject()
{
var xmlHttp=null;
try
  {
  // Firefox, Opera 8.0+, Safari
  xmlHttp=new XMLHttpRequest();
  }
catch (e)
  {
  // Internet Explorer
  try
    {
    xmlHttp=new ActiveXObject("Msxml2.XMLHTTP");
    }
  catch (e)
    {
    xmlHttp=new ActiveXObject("Microsoft.XMLHTTP");
    }
  }
return xmlHttp;
}

function Toggle3(item) {
   obj=document.getElementById(item);
   objdate=document.getElementById("d" + item);

   GetDetails3(objdate.innerHTML);
   
   visible=(obj.style.display!="none")
   key=document.getElementById("x" + item);
   if (visible) {
     obj.style.display="none";
     key.innerHTML='[+]';
   } else {
      obj.style.display="block";
      key.innerHTML='[-]';
   }
}

var xmlHttp

function GetDetails3(str)
{ 
xmlHttp=GetXmlHttpObject();

if (xmlHttp==null)
  {
  alert ("Sorry, your browser does not support this function.");
  return;
  }
document.body.style.cursor='wait';        
var url="gosdb-getmatchdetails1.asp";
url=url+"?q="+str;
url=url+"&sid="+Math.random();
xmlHttp.onreadystatechange=stateChanged;
xmlHttp.open("GET",url,true);
xmlHttp.send(null);
document.body.style.cursor='auto';   
}

function Toggle(item) {
   obj1=document.getElementById("contribution_abbrev" + item);
   obj2=document.getElementById("contribution_full" + item);
   visible=(obj1.style.display!="none");
   if (visible) {
     obj1.style.display="none";
     obj2.style.display="block";
   } else {
      obj1.style.display="block";
      obj2.style.display = "none";
   }
}
//-->

</script>

</head>

<body>
<!--#include file="top_code.htm"-->

<%

dim conn, sql, rs, fs, outline, photoname, notes, penpic, i, j, forename, surname, initials, fullname, akaname, akasurname, dob, lastgameyear, primephoto, primephotono, uniqueid
dim work1, work2, currentind, mod97num, abbrev_text, full_text, fulldisplayoption, fulltoggleoption, contributor, contributor_text, abbrev_len

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

	sql = "select player_id_spell1, surname, forename, initials, full_forenames, aka_forename, aka_surname, dob, last_game_year, prime_photo, penpic, notes "
	if Request.QueryString("status") = "preview" then sql = sql & ", penpic_pending "
	sql = sql & "from player  "
	sql = sql & "where player_id = " & playerid
	
	rs.open sql,conn,1,2	  
		
	if not IsNull(rs.Fields("forename")) then forename = trim(rs.Fields("forename"))
	if not IsNull(rs.Fields("surname")) then surname = trim(rs.Fields("surname"))
	if not IsNull(rs.Fields("initials")) then initials = trim(rs.Fields("initials")) 
	if not IsNull(rs.Fields("full_forenames")) then fullname = trim(rs.Fields("full_forenames")) & " " & trim(rs.Fields("surname"))
	if not IsNull(rs.Fields("aka_forename")) then akaname = trim(rs.Fields("aka_forename"))
	if not IsNull(rs.Fields("aka_surname")) then akasurname = trim(rs.Fields("aka_surname"))
	if not IsNull(rs.Fields("dob")) then dob = trim(rs.Fields("dob"))
	if not IsNull(rs.Fields("penpic")) then 
		penpic = replace(rs.Fields("penpic"),"|p|","</p><p style=""margin:4 12 0 0; text-align:left; line-height:1.3;"">")
		penpic = replace(penpic,"Footnote:","<span style=""font-style:italic"">Footnote:</span>")
	end if
	if Request.QueryString("status") = "preview" then
		penpic = replace("<span style=""color:red;"">New pen-picture review:</span>|p|" & rs.Fields("penpic_pending"),"|p|","</p><p style=""margin:4 12 0 0; text-align:left; line-height:1.3;"">")
		penpic = replace(penpic,"Footnote:","<span style=""font-style:italic"">Footnote:</span>")
	end if
	if not IsNull(rs.Fields("notes")) then notes = replace(rs.Fields("notes"),"|p|","</p><p style=""margin:4 12 0 0; text-align:left; line-height:1.3;"">")
		
	if not IsNull(rs.Fields("prime_photo")) then
		primephotono = rs.Fields("prime_photo")
		primephoto = "_" & rs.Fields("prime_photo")
	  else 
		primephotono = ""
		primephoto = ""
	end if
	
	lastgameyear = rs.Fields("last_game_year")
		
	if len(rs.Fields("player_id_spell1")) < 4 then 
		photoname = right("00" & rs.Fields("player_id_spell1"),3)
	  else
	  	photoname = rs.Fields("player_id_spell1")
	end if
	
	rs.close
%>

  
  <div id="container">
        
	<div id="left">

	<p style="text-align: center; margin-top:0; margin-bottom:0; padding-top:4; padding-left:8;" >
	<a href="gosdb.asp"><font color="#404040"> 
    <img border="0" src="images/gosdb-small.jpg"></font></a><font color="#404040">
    </p>

<%
	outline = ""

	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	if fs.FileExists(Server.MapPath("gosdb/photos/players/" & photoname & primephoto & ".jpg")) then
        outline = outline & "<img style=""margin-top:0px;"" border=""0"" src=""gosdb/photos/players/" & photoname & primephoto & ".jpg"" name=""photo"">"
      else 
      	outline = outline & "<img style=""margin-top:0px;"" border=""0"" src=""gosdb/photos/players/nophoto.jpg"">"
    end if
    
    outline = outline & "<p style=""margin:0 0 0 8"">"
   
    for i = 1 to 15
				
 		if fs.FileExists(Server.MapPath("gosdb/photos/players/" & photoname & "_" & i & ".jpg")) then	
 			if i = 1 then
 			    if fs.FileExists(Server.MapPath("gosdb/photos/players/" & photoname & "_" & i & ".jpg")) then	
 					if primephotono = "" then 
 						outline = outline & "<a href=""#"" onmouseover=""roll_over('photo','gosdb/photos/players/" & photoname & ".jpg" & "')"">"
 						outline = outline & "<img border=""0"" src=""gosdb/photos/players/" & photoname& ".jpg"" width=""58"" height=""77"" hspace=""0"">"
 						outline = outline & "</a>"
 						j = 1 
	 				  else
	 					outline = outline & "<a href=""#"" onmouseover=""roll_over('photo','gosdb/photos/players/" & photoname & "_" & primephotono & ".jpg" & "')"">"
	 					outline = outline & "<img border=""0"" src=""gosdb/photos/players/" & photoname & "_" & primephotono & ".jpg"" width=""58"" height=""77"" hspace=""0"">"
	 					outline = outline & "</a>" 
 						outline = outline & "<a href=""#"" onmouseover=""roll_over('photo','gosdb/photos/players/" & photoname & ".jpg" & "')"">"
 						outline = outline & "<img border=""0"" src=""gosdb/photos/players/" & photoname& ".jpg"" width=""58"" height=""77"" hspace=""0"">"
 						outline = outline & "</a>"
 						j = 2 
 					end if
				end if	
			end if
			if i <> primephotono then
	 			outline = outline & "<a href=""#"" onmouseover=""roll_over('photo','gosdb/photos/players/" & photoname & "_" & i & ".jpg" & "')"">"
	 			outline = outline & "<img border=""0"" src=""gosdb/photos/players/" & photoname & "_" & i & ".jpg"" width=""58"" height=""77"" hspace=""0"">"
				outline = outline & "</a>"
				j = j + 1
			end if
			if j = 4 or j = 8 then outline = outline & "<p style=""margin:0 0 0 8"">"
		  else exit for
 		end if
 	next
  	if lastgameyear > 1984 then outline = outline & "<p style=""font-size: 10px; margin:0px 12px 0 12px;  text-align:justify"">Thanks to Dave Rowntree for many of the player images after 1984.</p>"
 	outline = outline & "<p style=""font-size: 10px; margin:12px 12px 12px 12px; text-align:justify""><b>Can you help?</b> This page is the result of the best endeavours of all concerned. If you spot a mistake or know of facts to add, or have a better photo, please get in touch using 'Contact Us' (top, right).</p>"

 	response.write(outline)
%>
	</div>
    
    <div id="right">
    <div id="right-left">
    <div id="centerpanel_manager"></div>
    <div id="centerpanel_player">
<%  
	outline = "" 
	
	if from = "squad" then
		outline = outline & "<p style=""margin:9 0 0 0""><a href=""squad.asp""><u>Back to Current Squad Page</u></a></p>"	
	  elseif from = "appear" then
		outline = outline & "<p style=""margin:9 0 0 0""><a href=""progressappears.asp""><u>Back to Season Appearance Chart</u></a></p>"	
	  else
		outline = outline & "<p style=""margin:9 0 0 0""><a href=""gosdb.asp""><u>Back to GoS-DB Hub</u></a>&nbsp;&nbsp; <a href=""gosdb-players1.asp?scope=" & scope & """><u>Find Another Player</u></a></p>"	
	end if
	
	outline = outline & "<div style=""margin: 12 0 12 0;"">"

	outline = outline & "<p style=""margin: 9 0 9 0; font-size:18px; color: #0f6e3c;""><b>"
	if forename = "" then
		outline = outline & initials & " " & UCase(surname)
	  else
	  	outline = outline & UCase(forename) & " " & UCase(surname)
	end if
	
	if akasurname > "" then outline = outline & " (aka " & Ucase(akasurname) & ")"
	
	outline = outline & "</b></p>"
	
	if notes > "" then
		outline = outline & "<div style=""margin: 0 0 12 0; padding: 6 6 6 9; border: 1px dashed #000000"">"
		outline = outline & "<p style=""margin:0 0 0 0; text-align:left; line-height:1.3;"">Note: " & notes & "</p>"
		outline = outline & "</div>"
	end if
	
	if fullname > "" or akaname > "" then
		outline = outline & "<p style=""margin: 0 0 4 0;"">"
		if fullname > "" then outline = outline & "<span style=""color:#0f6e3c; font-weight:bold;"">Full Name: </span>" & fullname
		if akaname > "" then outline = outline & " (also known as " & akaname & ")"
		outline = outline & "</p>"
	end if

	if dob > "" then outline = outline & "<p class=""season"" style=""margin: 0 0 4 0;""><span style=""color:#0f6e3c; font-weight:bold;"">Born: </span>" & FormatDateTime(dob,1)	

	currentind = ""
		
	sql = "select spell, came_from, went_to, last_game_year, min(date) as firstdate, max(date) as lastdate " 
	sql = sql & "from player a left outer join match_player b on a.player_id = b.player_id "
	sql = sql & "where player_id_spell1 = " & playerid
	sql = sql & "group by spell, came_from, went_to, last_game_year " 
	sql = sql & "order by spell "
		
	rs.open sql,conn,1,2
	
	Do While Not rs.EOF
		outline = outline & "<p style=""margin: 0 0 4 0;"">"
		if rs.RecordCount > 1 then outline = outline & rs.Fields("spell") & ". "
		outline = outline & "<span style=""color:#0f6e3c; font-weight:bold;"">Came from: </span>" & rs.Fields("came_from") & "&nbsp;&nbsp;&nbsp;<span style=""color:#0f6e3c; font-weight:bold;"">Went to: </span>" & rs.Fields("went_to") & "</p>"
		outline = outline & "<p style=""margin: 0 0 4 0;"">"
		if rs.RecordCount > 1 then outline = outline & "<span style=""color:white"">" & rs.Fields("spell") & ". </span>"
		if not isnull(rs.Fields("firstdate")) then outline = outline & "<span style=""color:#0f6e3c; font-weight:bold;"">First game: </span>" & FormatDateTime(rs.Fields("firstdate"),1) & "&nbsp;&nbsp;&nbsp;<span style=""color:#0f6e3c; font-weight:bold;"">Last game: </span>" & FormatDateTime(rs.Fields("lastdate"),1) & "</p>"
		if rs.Fields("last_game_year") = 9999 then currentind = "Y"
		rs.MoveNext
	Loop
	rs.close
	
	sql = "with cte as ( "
	sql = sql & "select count(*) as starts, 0 as subs, 0 as goals "
	sql = sql & "from player a join match_player b on a.player_id = b.player_id "
	sql = sql & "where player_id_spell1 = " & playerid
	sql = sql & "  and startpos > 0 "
	sql = sql & "union all "
	sql = sql & "select 0, count(*), 0 "
	sql = sql & "from player a join match_player b on a.player_id = b.player_id "
	sql = sql & "where player_id_spell1 = " & playerid
	sql = sql & "  and startpos = 0 "
	sql = sql & "union all "
	sql = sql & "select 0, 0, count(*) "
	sql = sql & "from player a join match_goal b on a.player_id = b.player_id "
	sql = sql & "where player_id_spell1 = " & playerid
	sql = sql & ") "
	sql = sql & "select sum(starts) as starts, sum(subs) as subs, sum(goals) as goals "
	sql = sql & "from cte " 
	
	rs.open sql,conn,1,2
	
	outline = outline & "<p style=""margin: 0 0 4 0;""><span style=""color:#0f6e3c; font-weight:bold;"">Appearances: </span>" & rs.Fields("starts") + rs.Fields("subs")
	outline = outline & " (" & rs.Fields("starts") & "/" & rs.Fields("subs") & ")" 
	outline = outline & "&nbsp;&nbsp;&nbsp;<span style=""color:#0f6e3c; font-weight:bold;"">Goals: </span>" & rs.Fields("goals") & "</p>"
	
	rs.close
		
	if penpic > "" then outline = outline & "<p style=""margin:6 12 6 0; text-align:left; line-height:1.3;"">" & penpic & "</p>"
	
	
		'Prepare a unique id for a user contribution
		sql = "select cast(replace(NEWID(),'-','') as char(32)) as uniqueid "
		rs.open sql,conn,1,2
		uniqueid = rs.Fields("uniqueid")
		rs.close
	
		'Prepare a Steve's version of a modulus 97 number as a means validating the url
		Randomize	
		work1 = int(((1030927 - 103093) * rnd) + 103093) * 97 	'calculates a random multiple of 97 with 8 digits
		work2 = int((rnd*95) + 1)								'random number between 1 and 96
		mod97num = CStr(work1 + work2) & right("0" & work2,2)
		
		outline = outline & "<div id=""contrib_head"">"			
		outline = outline & "<p style=""margin: 4px 0 3px 0; color: #0f6e3c;"";""><b>YOUR CONTRIBUTION</b></p>"
		
		if currentind = "" then	
			outline = outline & "<p style=""margin: 0;"">If you can add to this profile, perhaps with special memories, a favourite story or the results of your original research, please "
			outline = outline & "<a href=""gosdb-players2-contribute1.asp?parm=" & uniqueid & right("000" & playerid,4) & mod97num & """><u>contribute here</u></a>.</p>"
		  else
			outline = outline & "<p style=""margin: 0;"">Sorry, new contributions are not accepted until the player leaves the club.</p>"
		end if	
		
		sql = "select contributor, location, datetime_added, text "
		sql = sql & "from player_contribution "
		sql = sql & "where player_id = " & playerid
		sql = sql & "  and isnull(rejected,'N')  <> 'Y' "
		sql = sql & "order by datetime_added "
		rs.open sql,conn,1,2
		
		i = 1
		
		Do While Not rs.EOF
		
			fulldisplayoption = "display:block;"
			fulltoggleoption = ""
			
			'check if contributor's name includes a ( and break into two if so
			contributor = split(trim(rs.Fields("contributor")),"(")
			if uBound(contributor) = 0 then
				contributor_text = contributor(0) & "</span>"
			  else
				contributor_text = contributor(0) & "</span> (" & contributor(1)
			end if 
			
			if rs.RecordCount = 1 and len(rs.Fields("text")) < 2000 then
				
				' if single contribution and not too long, bypass a short version and display it in full
				
			  elseif len(rs.Fields("text")) > 600 then		'if not too short, create the abbreviated option
			  
			  	abbrev_len = 600
			  	if rs.RecordCount = 1 then abbrev_len = 1000 
			
			  	abbrev_text = left(left(rs.Fields("text"),abbrev_len),instrrev(left(rs.Fields("text"),abbrev_len)," "))
			  	if instrrev(abbrev_text,"|p|") > abbrev_len-100 then abbrev_text = left(abbrev_text,instrrev(abbrev_text,"|p|")-1)		'reduce even more if new paragraph is within 100 of the end 
				abbrev_text = replace(abbrev_text,"|p|","</p><p style=""margin: 4px 0;"">") 
				abbrev_text = replace(abbrev_text,"[b]","<span style=""font-weight: bold"">") 
				abbrev_text = replace(abbrev_text,"[i]","<span style=""font-style: italic"">")
				abbrev_text = replace(abbrev_text,"[u]","<span style=""text-decoration: underline"">")
				abbrev_text = replace(abbrev_text,"[in]","<div style=""margin:4px 18px"">")
				abbrev_text = replace(abbrev_text,"[/b]","</span>")
				abbrev_text = replace(abbrev_text,"[/i]","</span>")
				abbrev_text = replace(abbrev_text,"[/u]","</span>")
				abbrev_text = replace(abbrev_text,"[/in]","</div>")
				
				outline = outline & "<div id=""contribution_abbrev" & i & """ style=""display:block; line-height:1.3;"">"
				outline = outline & "<p style=""margin: 9px 0 0;""><span style=""font-weight: bold; color: #202020;"">" 
				
				if left((contributor_text),13) = "Brian Knight*" then
					outline = outline & "By " & contributor_text
				  else	
					outline = outline & "From " & contributor_text
				end if	
				if trim(rs.Fields("location")) > " " then 
					if left(Ucase(trim(rs.Fields("location"))),5) = "NEAR " then
						outline = outline & " " & trim(rs.Fields("location")) 					
					  else
						outline = outline & " in " & trim(rs.Fields("location"))
					end if	 
				end if		
				outline = outline & " on " & FormatDateTime(left(rs.Fields("datetime_added"),10),2) & " ...</p>"
				outline = outline & "<p style=""margin: 3px 0;"">" & abbrev_text & " ..."
				outline = outline & " &nbsp;<a href=""javascript:Toggle('" & i & "')""><u>More</u></a>" 
				outline = outline & "</p></div>"
				
				fulldisplayoption = "display:none;"
				fulltoggleoption = "</p><p style=""margin: 4px 0;""><a href=""javascript:Toggle('" & i & "')""><u>Close above</u></a>"
			
			end if 
			
			full_text = replace(rs.Fields("text"),"|p|","</p><p style=""margin: 4px 0;"">") 
			full_text = replace(full_text,"[b]","<span style=""font-weight: bold"">") 
			full_text = replace(full_text,"[i]","<span style=""font-style: italic"">")
			full_text = replace(full_text,"[u]","<span style=""text-decoration: underline"">")
			full_text = replace(full_text,"[in]","<div style=""margin:4px 18px"">")
			full_text = replace(full_text,"[/b]","</span>")
			full_text = replace(full_text,"[/i]","</span>")
			full_text = replace(full_text,"[/u]","</span>")
			full_text = replace(full_text,"[/in]","</div>")	
				 		
			outline = outline & "<div id=""contribution_full" & i & """ style=""" & fulldisplayoption & " line-height: 1.3;"">"
			outline = outline & "<p style=""margin: 9px 0 0;""><span style=""font-weight: bold; color: #202020;"">" 

			if left((contributor_text),13) = "Brian Knight*" then
				outline = outline & "By " & contributor_text
			  else	
				outline = outline & "From " & contributor_text
			end if	
			if trim(rs.Fields("location")) > " " then 
				if left(Ucase(trim(rs.Fields("location"))),5) = "NEAR " then
					outline = outline & " " & trim(rs.Fields("location")) 					
				  else
					outline = outline & " in " & trim(rs.Fields("location"))
				end if	 
			end if	
			outline = outline & " on " & FormatDateTime(left(rs.Fields("datetime_added"),10),2) & " ...</p>"
			outline = outline & "<p style=""margin: 3px 0;"">" & full_text
			outline = outline & fulltoggleoption 
			outline = outline & "</p></div>"
  			i = i + 1
			rs.MoveNext
		
		Loop
		rs.close
		
		outline = outline & "</div>"
		
	outline = outline & "</div>"	

	
	outline = outline & "<p style=""margin: 12 0 4 0; color: #0f6e3c;""><b>APPEARANCE DETAILS</b>&nbsp;&nbsp;[<a href=""gosdb-players0.asp""><u>reselect competitions</u></a>]</p>"	
	outline = outline & "<p style=""margin: 4 0 2 0; color: #cc3300;"">"
    
    if scope = "1,2,3,4,5,6,7" then
    
    	outline = outline & "The details below reflect appearances in all first-team competitions."
    	
      else
		
		outline = outline & "<b>The details below reflect appearances in the following selected competitions: "

		sql = "select distinct LFC, compcat, compcatname " 
		sql = sql & "from competition " 
		sql = sql & "where compcat in (" & scope & ") "
		sql = sql & "order by compcat "
		rs.open sql,conn,1,2
	
		Do While Not rs.EOF
	  	outline = outline & rs.Fields("compcatname") & ", "
	  	rs.MoveNext
		Loop
		
		if rs.RecordCount = 0 then 
			outline = outline & "None</b>"  	'remove last comma and space 
			else outline = left(outline,len(outline)-2) & "</b>"  	'remove last comma and space
		end if
		rs.close
		
		outline = outline & "</p><p style=""margin: 2 0 2 0; color: #cc3300;"">Excluded: "

		sql = "select distinct LFC, compcat, compcatname " 
		sql = sql & "from competition " 
		sql = sql & "where not compcat in (" & scope & ") "
		sql = sql & "order by compcat "
		rs.open sql,conn,1,2
	
		Do While Not rs.EOF
	  	outline = outline & rs.Fields("compcatname") & ", "
	  	rs.MoveNext
		Loop
		rs.close 
	
		outline = left(outline,len(outline)-2)  	'remove last comma and space 
	
	end if
	
	conn.close
	
	outline = outline & "</p>"
	
	response.write(outline)
%>
    	
    <div id="ajaxplayerdetails">
    </div>
	
	</div>	<!-- end of centerpanel_player -->
	</div>	<!-- end of right-left -->

		
	<div id="right-right">
   	<div id="ajaxplayerlist">
	</div> 
    
    </div>	<!-- end of right-right -->
    </div>	<!-- end of right -->



	<div id="bottom">
	<p style="margin-left:18; margin-right:12; margin-top:24; margin-bottom:0" align="justify"><span style="font-size: 10px">
    I'm very grateful to many who have helped write GoS-DB's player 
    pen-pictures, and to Dave Rowntree, the PAFC Media Team and Colin Parsons for their help with photos. 
    Thanks also to staff at the National Football Museum, the Scottish Football 
    Museum and ScotlandsPeople for their valuable assistance.</span><p style="margin-left:18; margin-right:12; margin-top:6; margin-bottom:0" align="justify">
    <span style="font-size: 10px">The following publications have been 
    particularly valuable in the research of pen-pictures: Plymouth Argyle, A 
    Complete Record 1903-1989 (Brian Knight, ISBN 0-907969-40-2); Plymouth 
    Argyle, 101 Golden Greats (Andy Riddle, ISBN 1-874287-47-3); Football League 
    Players' Records 1888-1939 (Michael Joyce, ISBN 1-899468-67-6); Football 
    League Players' Records 1946-1988 (Barry Hugman, ISBN 1-85443-020-3) and 
    Plymouth Argyle Football Club Handbooks.</span></td>

	</div>	<!-- end of bottom -->
	</div>	<!-- end of container -->

	
<!--#include file="base_code.htm"-->

</body>

</html>