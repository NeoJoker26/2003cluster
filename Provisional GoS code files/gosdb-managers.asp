<%@ Language=VBScript %>
<% Option Explicit %>
<%
Dim id1, id2 
id1 = Request.QueryString("id1")
id2 = Request.QueryString("id2")
%>

<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<title>GoS-DB Managers</title>
<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
<link rel="stylesheet" type="text/css" href="gos2.css">
<style>
<!--
.hide {display:none;}
.rowhlt {background-color: #e3eee3;}
.hover {color: #000000; background-color: #c8e0c7 !important; cursor: pointer;}
.close {cursor: pointer;}
.grey {background-color: #e0e0e0;}
.green {background-color: #e3eee3;}
.black {background-color: #202020; color: #ffffff;}

.leftpad { 
	padding-left:5px !important;
}

#managers {
	border-collapse: collapse; 
	margin: 12px 0;
}
#managers td {
	border: 1px solid #c0c0c0;
	padding: 3px 1px 3px 1px;
	text-align: right; 
	vertical-align: top;
}
#managers th {
	border: 1px solid #c0c0c0;
	text-align: center;
	padding: 5px 1px 5px 1px;
	vertical-align: bottom;
}

#managers td p {
	text-align: left;
}

#manageres ol {
	font-size:11px;
}

#managers .manager_name {
	text-align: left;
	white-space: nowrap;
	padding-left: 2px; padding-right: 2px;
}

.manager_details {
	vertical-align: text-bottom; 
	margin:0 1px 0 0; 
	padding:0
}

.playerlisthead {
	font-weight: bold;
	margin: 9px 0 4px;
}

.playerlist {
	margin: 4px 0 9px;
}

.date {
	white-space: nowrap;
	}	

.button {
    padding: 4px 8px 5px; 
    margin-right: 8px;
    margin-bottom:10px;
    font-family: verdana, sans-serif;
    display: inline-block;
    white-space: nowrap;
    font-size:11px;
    position:relative;
    outline: none;
    overflow: visible;
    cursor: pointer;
    border-radius: 4px;
    border: 1px solid #808080;
    color: #000000 !important; 
    background: linear-gradient(#e0f0e0,#d0e0d0);  
}

.button:hover {
    opacity:0.70; 
    filter:alpha(opacity=70);
    border: 1px solid #202020;        
}

.ui-dialog {
	font-family: verdana,arial,helvetica,sans-serif;
	font-size: 12px;
	background: #e3eee3;
}

.ui-dialog-content {
	text-align: center;
	line-height: 1.8;
}
-->
</style>

<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script>
$(document).ready(function(){
  	$(function() {
    	$("#dialog").dialog({position: {my: "center top", at: "center top", of: "#managers"} });
    });
	$("#managers tr:not('.footnote,.header')").mouseenter(function() {
	    	$(this).addClass("rowhlt");
	});
	$("#managers tr:not('.footnote,.header')").mouseleave(function() {
    	$(this).removeClass("rowhlt");
	});

	$(".manager_name_hover").hover(function() {
    	$(this).toggleClass("hover");
	});

	$(".manager_name_hover").click(function(){
		var cellnamea = $(this).attr('id');
		if (cellnamea == 'dialog') {
			cellnamea = $(this).attr('class');
			cellnamea = cellnamea.split(' ')[0];
			$("#dialog").hide();
			}
		var temp = cellnamea.substring(6)
		var managerids = temp.split("-")  
		var rownamea = cellnamea.replace("cell-a","row-a");
		var rownameb = rownamea.replace("row-a","row-b");  
		var cellnameb = rownameb.replace("row-b","cell-b");
		var ajaxparm = "id1=" + managerids[0] + "&id2=" + managerids[1]
		if ($("#" + rownameb).is(':hidden')) {
				$("#" + cellnameb).html('').load('gosdb-getmanagertext.asp?' + ajaxparm,function(response, status){
		  		if (status == 'success')
				$("#" + rownameb).show('slow');
				$("#" + rownamea).addClass("black");
				$('html, body').animate({scrollTop: $("#" + rownamea).offset().top}, 1000);
        		});
        	};
	});
	
	$("#managers").on("click",".close", function(){
		var cellnameb = $(this).parent().attr('id');
		var rownameb = cellnameb.replace("cell-b","row-b");
		var rownamea = rownameb.replace("row-b","row-a");  
		$("#" + rownameb).hide('slow');
		$("#" + rownamea).removeClass("black");
	});
	
	$('#selectvenue').change(function() {
        this.form.submit();
    });
	
	<%
	if id1 > "" then
		if id2 = "" then id2 = 999
		response.write("$('#cell-a" & id1 & "-" & id2 & "-1" & "').trigger('click');")
	end if
	%>

});
</script>

</head>

<body>
<!--#include file="top_code.htm"-->
<%
Dim conn, sql, rs, rs1, n, sort, outline, daysic, orderby_text, restrictions, heading1, heading2, venue, venueclause, selected_venue(2)
Dim date1, date2, fromdates, todates, managers, focusname, dialoghold
Dim noctk, nospell, caretaker_clause, caretaker_option1, caretaker_option2, caretaker_value, dbview, spell_option1, spell_option2, spell_value
Dim rowclass, rowid, footnote_no, footnotes, footnote
  
sort = Request.QueryString("sort")
if sort = "" then sort = Request.Form("sort")
if sort = "" then sort = 1

noctk = Request.QueryString("noctk")
if noctk = "" then noctk = Request.Form("noctk")
if noctk = "" then noctk = 0

nospell = Request.QueryString("nospell")
if nospell = "" then nospell = Request.Form("nospell")
if nospell = "" then nospell = 0 	

venue = Request.QueryString("venue")
if venue = "" then venue = Request.Form("venue")	

focusname = Request.QueryString("focus")

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs1 = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%
%>
<div style="margin:0px auto; width:980px">
  <table>
    <tr>
      <td width="260" valign="top" style="text-align: left">
		<div style="width:260;">
		<p style="text-align: center; margin-top:0; margin-bottom:3">
		<a href="gosdb.asp"><font color="#404040"><img border="0" src="images/gosdb-small.jpg" align="left"></font></a><font 
		color="#404040"> 
		<b><font style="font-size: 15px">Search by<br>
		</font></b><span style="font-size: 15px"><b>Manager</b></span></font><p style="text-align: center; margin-top:0; margin-bottom:0">
		<b>
		<a href="gosdb.asp">Back to<br>GoS-DB Hub</a> </b>
		</div>
      </td>
      <td align="center" valign="top" style="text-align: center" width="460">
      <p style="margin-top: 9px; margin-bottom: 9px">
      <span style="font-size: 18px"><font color="#006E32">
      MANAGERS</font></p>
      <p style="font-weight:bold; margin-top:9px; margin-bottom:9px">Choose to exclude/include caretaker managers</br>and to combine/separate multiple spells ...</p>
      
<%
 select case sort
	case 1
		orderby_text = " sortdate "
		heading2 = heading2 & "Ordered by appointment date"
	case 2 
		orderby_text = " daysic desc, sortdate "
		heading2 = heading2 & "Ordered by days in the job"
	case 3 
		orderby_text = " P desc, sortdate "
		heading2 = heading2 & "Ordered by league games played" 
	case 4
		orderby_text = " W desc, sortdate "
		heading2 = heading2 & "Ordered by wins"
	case 5
		orderby_text = " D desc, sortdate "
		heading2 = heading2 & "Ordered by draws"
	case 6
		orderby_text = " L desc, sortdate "
		heading2 = heading2 & "Ordered by defeats"
	case 7
		orderby_text = " F desc, sortdate "
		heading2 = heading2 & "Ordered by goals-for"
	case 8
		orderby_text = " A desc, sortdate "
		heading2 = heading2 & "Ordered by goals-against"
	case 9
		orderby_text = " PO desc, sortdate "
		heading2 = heading2 & "Ordered by points (3 for a win in all cases)"
	case 10
		orderby_text = " Wperc desc, sortdate "
		heading2 = heading2 & "Ordered by percentage wins" 
	case 11
		orderby_text = " Dperc desc, sortdate "
		heading2 = heading2 & "Ordered by percentage draws"
	case 12
		orderby_text = " Lperc desc, sortdate "
		heading2 = heading2 & "Ordered by percentage defeats"
	case 13
		orderby_text = " Favg desc, sortdate "
		heading2 = heading2 & "Ordered by average goals-for"
	case 14
		orderby_text = " Aavg desc, sortdate "
		heading2 = heading2 & "Ordered by average goals-against"
	case 15
		orderby_text = " POavg desc, sortdate "
		heading2 = heading2 & "Ordered by average points (3 for a win in all cases)"
	case 16
		orderby_text = " CWperc desc, sortdate "
		heading2 = heading2 & "Ordered by percentage wins in cups"				
 end select	

  if noctk = 1 then
  	caretaker_option1 = "Include"
  	caretaker_option2 = "excluded"
   	caretaker_value = 0
   else
   	caretaker_option1 = "Exclude"
   	caretaker_option2 = "included"
    caretaker_value = 1
  end if
  if nospell = 1 then
  	spell_option1 = "Separate"
  	spell_option2 = "combined"
  	spell_value = 0
   else
   	spell_option1 = "Combine"
   	spell_option2 = "separated"
    spell_value = 1
  end if	
      
  %>

  <a class="button" href="gosdb-managers.asp?sort=<%response.write(sort)%>&noctk=<%response.write(caretaker_value)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
  <%response.write(caretaker_option1)%> Caretakers</a>
  <a class="button" href="gosdb-managers.asp?sort=<%response.write(sort)%>&noctk=<%response.write(noctk)%>&nospell=<%response.write(spell_value)%>&venue=<%response.write(venue)%>">
  <%response.write(spell_option1)%> Spells</a>
  
  <form style="padding: 0; margin: 0;" action="gosdb-managers.asp" method="post" name="form1">
  <%  
	select case venue
	case "home"
		selected_venue(1) = "selected"
		venueclause = "where homeaway = 'H' "
		heading1 = "Home matches only, "
	case "away"
		selected_venue(2) = "selected"
		venueclause = "where homeaway = 'A' "
		heading1 = "Away matches only, "
	case else
		selected_venue(0) = "selected"
		venueclause = ""
		heading1 = "All matches, "
 	end select
 
 	response.write("<select name=""venue"" id=""selectvenue"" style=""font-size: 11px; padding: 2px 4px;"">")
 	response.write("<option value=""all"" " & selected_venue(0) & ">For All Matches</option>")
 	response.write("<option value=""home"" " & selected_venue(1) & ">Home Games Only</option>") 
 	response.write("<option value=""away"" " & selected_venue(2) & ">Away Games Only</option>")
	response.write("</select>")
	response.write("<input type=""hidden"" name = ""noctk"" value=""" & noctk & """>")
	response.write("<input type=""hidden"" name = ""nospell"" value=""" & nospell & """>")
	response.write("<input type=""hidden"" name = ""sort"" value=""" & sort & """>")
  %>
  </form>
  	
  <% response.write("<p class=""style1boldgreen"" style=""font-size:12px"">" & Ucase(heading1 & heading2) & "</p>") %>
    
  </td>
      <td width="260" valign="top" style="text-align: center">
      <p style="margin-bottom:3pt; margin-right:3; text-align:justify; margin-top:3">Just a mass 
      of numbers? Each manager's name leads to a full profile, and the page 
      comes into its own when you use the 'sort' buttons to bring out the best 
      and worst of any column. </p>
      </td>
    </tr>
  </table>
  
  <table id="managers">

    <tr class="header">
      <th style="border: 0px none;" rowspan="4"></th>
      <th style="color:#006E32" colspan="5">
      <% response.write("Caretakers " & caretaker_option2 & ", multiple spells " & spell_option2) %>
      </th>
      <th style="border-top:0px none; border-bottom:0px none;">&nbsp;</th>
      <th colspan="13"><b>Leagues</b></th>
      <th style="border-top:0px none; border-bottom:0px none;">&nbsp;</th>
      <th colspan="7"><b>Cups</b></th>
    </tr>
    <tr class="header">
      <th style="border-bottom:0px none; padding-bottom:0; padding-left:2px; padding-right:2px" rowspan="2">Foot<br>note</th> 
      <th style="border-bottom:0px none; padding-bottom:0" rowspan="2"><b>Manager</b><%if noctk = 0 then response.write("</br><span style=""font-size:10px; font-weight:bold"">(caretakers shaded grey)</span>")%></th>
      <th style="border-bottom:0px none; padding-bottom:0" class="textleft" rowspan="2"><b>Appointed</b></th>
      <th style="border-bottom:0px none; padding-bottom:0" class="textleft" rowspan="2"><b>Departed</b></th>
      <th style="border-bottom:0px none; padding-bottom:0" rowspan="2"><b>Days</b></th>    
      	
      <th style="border-top:0px none; border-bottom:0px none;" rowspan="2">&nbsp;</th>
      
      <th style="border-bottom:0px none; padding-bottom:0" rowspan="2"><b>P</b></th>
      <th style="border-bottom:0px none; padding-bottom:0" rowspan="2"><b>W</b></th>
      <th style="border-bottom:0px none; padding-bottom:0" rowspan="2"><b>D</b></th>
      <th style="border-bottom:0px none; padding-bottom:0" rowspan="2"><b>L</b></th>
      <th style="border-bottom:0px none; padding-bottom:0" rowspan="2"><b>F</b></th>
      <th style="border-bottom:0px none; padding-bottom:0" rowspan="2"><b>A</b></th>
      <th style="border-bottom:0px none; padding-bottom:0" rowspan="2"><b>Pts<sup>1</sup></b>  
      
      <th colspan="3">Percent</th>
      <th colspan="3">Average</th>  
      
      <th style="border-top:0px none; border-bottom:0px none;" rowspan="2">&nbsp;</th>
        
      <th style="border-bottom:0px none; padding-bottom:0" rowspan="2"><b>P</b></th>
      <th style="border-bottom:0px none; padding-bottom:0" rowspan="2"><b>W</b></th>
      <th style="border-bottom:0px none; padding-bottom:0" rowspan="2"><b>D</b></th>
      <th style="border-bottom:0px none; padding-bottom:0" rowspan="2"><b>L</b></th>
      <th style="border-bottom:0px none; padding-bottom:0" rowspan="2"><b>F</b></th>
      <th style="border-bottom:0px none; padding-bottom:0" rowspan="2"><b>A</b></th>
      <th style="border-bottom:0px none; padding-bottom:0" rowspan="2"><b>%<br>W</b></th>
    </tr>

    <tr class="header">   
      <th style="border-bottom:0px none; padding-bottom:0"><b>W</b></th>
      <th style="border-bottom:0px none; padding-bottom:0"><b>D</b></th>
      <th style="border-bottom:0px none; padding-bottom:0"><b>L</b></th>
      <th style="border-bottom:0px none; padding-bottom:0"><b>F</b></th>
      <th style="border-bottom:0px none; padding-bottom:0"><b>A</b></th>
      <th style="border-bottom:0px none; padding-bottom:0"><b>Pts<sup>1</sup></b></th>
    </tr>
    
    <tr class="header">
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px">&nbsp;</th>
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px"><img style="vertical-align: top; margin:0 2px 0 0; padding:0" src="images/more.png"><span style="font-weight:normal">Click on name for more</span></th>
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px"><a href="gosdb-managers.asp?sort=1&noctk=<%response.write(noctk)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
      	<img src="images/sort.gif" border="0"></a></th>
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px">&nbsp;</th>
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px"><a href="gosdb-managers.asp?sort=2&noctk=<%response.write(noctk)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
      	<img src="images/sort.gif" border="0"></a></th>
      <th style="border-top:0px none; border-bottom:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:6px">&nbsp;</th>
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px"><a href="gosdb-managers.asp?sort=3&noctk=<%response.write(noctk)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
      	<img src="images/sort.gif" border="0"></a></th>
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px"><a href="gosdb-managers.asp?sort=4&noctk=<%response.write(noctk)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
      	<img src="images/sort.gif" border="0"></a></th>
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px"><a href="gosdb-managers.asp?sort=5&noctk=<%response.write(noctk)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
      	<img src="images/sort.gif" border="0"></a></th>
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px"><a href="gosdb-managers.asp?sort=6&noctk=<%response.write(noctk)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
      	<img src="images/sort.gif" border="0"></a></th>
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px"><a href="gosdb-managers.asp?sort=7&noctk=<%response.write(noctk)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
      	<img src="images/sort.gif" border="0"></a></th>
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px"><a href="gosdb-managers.asp?sort=8&noctk=<%response.write(noctk)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
      	<img src="images/sort.gif" border="0"></a></th>
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px"><a href="gosdb-managers.asp?sort=9&noctk=<%response.write(noctk)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
       	<img src="images/sort.gif" border="0"></a></th>
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px"><a href="gosdb-managers.asp?sort=10&noctk=<%response.write(noctk)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
       	<img src="images/sort.gif" border="0"></a></th>
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px"><a href="gosdb-managers.asp?sort=11&noctk=<%response.write(noctk)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
       	<img src="images/sort.gif" border="0"></a></th>
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px"><a href="gosdb-managers.asp?sort=12&noctk=<%response.write(noctk)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
      	<img src="images/sort.gif" border="0"></a></th>
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px"><a href="gosdb-managers.asp?sort=13&noctk=<%response.write(noctk)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
      	<img src="images/sort.gif" border="0"></a></th>
      <th style="border-top:0px none; padding-left:0; padding-right:0; padding-top:0; padding-bottom:8px"><a href="gosdb-managers.asp?sort=14&noctk=<%response.write(noctk)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
      	<img src="images/sort.gif" border="0"></a></th>
      <th style="border-top:0px none; padding-bottom:8px"><a href="gosdb-managers.asp?sort=15&noctk=<%response.write(noctk)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
      	<img src="images/sort.gif" border="0"></a></th>
      	
      <th style="border-top:0px none; border-bottom:0px none;">&nbsp;</th>
      <th style="border-top:0px none;">&nbsp;</th>
      <th style="border-top:0px none;">&nbsp;</th>
      <th style="border-top:0px none;">&nbsp;</th>
      <th style="border-top:0px none;">&nbsp;</th>
      <th style="border-top:0px none;">&nbsp;</th>
      <th style="border-top:0px none;">&nbsp;</th>
      <th style="border-top:0px none; padding-bottom:8px"><a href="gosdb-managers.asp?sort=16&noctk=<%response.write(noctk)%>&nospell=<%response.write(nospell)%>&venue=<%response.write(venue)%>">
      	<img src="images/sort.gif" border="0"></a></th>

    </tr>

<%

if noctk = 1 then caretaker_clause  = "where caretaker is null "
      
sql = "WITH CTE1 AS ( "
sql = sql & "select managers, spell_no, caretaker, manager_id1, manager_id2, "
sql = sql & "sortdate, daysic, " 	
sql = sql & "sum(p) as P, sum(w) as W, sum(d) as D, sum(l) as L, sum(f) as F, sum(a) as A, sum(po) as PO, "
sql = sql & "sum(cp) as CP, sum(cw) as CW, sum(cd) as CD, sum(cl) as CL, sum(cf) as CF, sum(ca) as CA, penpiclen "
sql = sql & "from "
	sql = sql & "( "
	if nospell = 1 then
		sql = sql & "select a.managers, 0 as spell_no, a.caretaker, earliest_from_date as sortdate,  datediff(day, from_date, isnull(to_date,getdate())) + 1 as daysic, a.manager_id1, isnull(a.manager_id2, 999) as manager_id2, len(b.penpic1) as penpiclen, "
  	  else
		sql = sql & "select managers, spell_no, caretaker, from_date as sortdate, datediff(day, from_date, isnull(to_date,getdate())) as daysic, manager_id1, isnull(manager_id2, 999) as manager_id2, len(penpic1) as penpiclen, "
	end if 
	sql = sql & "case when LFC <> 'C' then 1 else 0 end as p, "
	sql = sql & "case when LFC <> 'C' and goalsfor > goalsagainst then 1 else 0 end as w, "
	sql = sql & "case when LFC <> 'C' and goalsfor = goalsagainst then 1 else 0 end as d, "
	sql = sql & "case when LFC <> 'C' and goalsfor < goalsagainst then 1 else 0 end as l, "
	sql = sql & "case when LFC <> 'C' then goalsfor else 0 end as f, " 
	sql = sql & "case when LFC <> 'C' then goalsagainst else 0 end as a, "  
	sql = sql & "case when LFC <> 'C' and goalsagainst = 0 then 1 else 0 end as cs, "  
	sql = sql & "case when LFC <> 'C' and goalsfor > goalsagainst then 3 when LFC <> 'C' and goalsfor = goalsagainst then 1 else 0 end as po, " 
	sql = sql & "case when LFC = 'C' then 1 else 0 end as cp, "
	sql = sql & "case when LFC = 'C' and goalsfor > goalsagainst then 1 else 0 end as cw, "
	sql = sql & "case when LFC = 'C' and goalsfor = goalsagainst then 1 else 0 end as cd, "
	sql = sql & "case when LFC = 'C' and goalsfor < goalsagainst then 1 else 0 end as cl, "
	sql = sql & "case when LFC = 'C' then goalsfor else 0 end as cf, "
	sql = sql & "case when LFC = 'C' then goalsagainst else 0 end as ca "
	if nospell = 1 then
		sql = sql & "from v_manager_horiz a join v_managerspell_horiz b on a.manager_id1 = b.manager_id1 and isnull(a.manager_id2, 999) = isnull(b.manager_id2, 999) "
  	  	sql = sql & "left outer join v_match_all on date between b.from_date and isnull(b.to_date,getdate()) "
  	  else
		sql = sql & "from v_managerspell_horiz "
		sql = sql & "left outer join v_match_all on date between from_date and isnull(to_date,getdate()) "
	end if
	sql = sql & venueclause
	sql = sql & ") as subsel "
sql = sql & caretaker_clause 
sql = sql & "group by managers, sortdate, daysic, caretaker, spell_no, manager_id1, manager_id2, penpiclen " 
sql = sql & "), "
sql = sql & "CTE2 as "
sql = sql & "( "
sql = sql & "select managers, spell_no, caretaker, "   
sql = sql & "sortdate, sum(daysic) as daysic, manager_id1, manager_id2, penpiclen, " 
sql = sql & "sum(p) as P, sum(w) as W, sum(d) as D, sum(l) as L, sum(f) as F, sum(a) as A, sum(po) as PO, "  
sql = sql & "sum(cp) as CP, sum(cw) as CW, sum(cd) as CD, sum(cl) as CL, sum(cf) as CF, sum(ca) as CA "
sql = sql & "from CTE1 "
sql = sql & "group by managers, spell_no, sortdate, caretaker, manager_id1, manager_id2, penpiclen "
sql = sql & "), "
sql = sql & "CTE3 as "
sql = sql & "( "
sql = sql & "select managers, spell_no, caretaker, sortdate, manager_id1, manager_id2, penpiclen, daysic, "
sql = sql & "P, W, D, L, F, A, PO, " 
sql = sql & "case when p > 0 then cast(round(100.0*w/p,1) as numeric(4,1)) else 0 end as Wperc, "
sql = sql & "case when p > 0 then cast(round(100.0*d/p,1) as numeric(4,1)) else 0 end as Dperc, "
sql = sql & "case when p > 0 then cast(round(100.0*l/p,1) as numeric(4,1)) else 0 end as Lperc, "
sql = sql & "case when p > 0 then cast(round(1.0*f/p,2) as numeric(3,2)) else 0 end as Favg, "
sql = sql & "case when p > 0 then cast(round(1.0*a/p,2) as numeric(3,2)) else 0 end as Aavg, "
sql = sql & "case when p > 0 then cast(round(1.0*po/p,2) as numeric(3,2)) else 0 end as POavg, "
sql = sql & "case when cp > 0 then cast(round(100.0*cw/cp,1) as numeric(4,1)) else 0 end as CWperc, "
sql = sql & "CP, CW, CD, CL, CF, CA "
sql = sql & "from CTE2 "
sql = sql & ") "
sql = sql & "select rank() over(order by " & orderby_text & ") as rowno, managers, spell_no, caretaker, sortdate, daysic, manager_id1, manager_id2, penpiclen,"   
sql = sql & "P, W, D, L, F, A, PO, "
sql = sql & "Wperc, Dperc, Lperc, Favg, Aavg, POavg, "
sql = sql & "CP, CW, CD, CL, CF, CA, CWperc "
sql = sql & "from CTE3 " 
sql = sql & "order by rowno "

rs.open sql,conn,1,2

outline = ""

Do While Not rs.EOF

	rowid = rs.Fields("manager_id1") & "-" & rs.Fields("manager_id2") & "-" & rs.Fields("spell_no") 

	rowclass = ""
	if rs.Fields("caretaker") = "Y" then rowclass = "grey"

	managers = replace(rs.Fields("managers"),"Larrieu & ","Larrieu & <br>")  
		
	daysic = rs.Fields("daysic")
	if len(daysic) > 3 then daysic = left(daysic,len(daysic)-3) & "," & right(daysic,3)
	 
	outline  = outline & "<tr id=""row-a" & rowid & """ class=""" & rowclass & """>"	

	outline  = outline & "<td style=""text-align: center"">" & rs.Fields("rowno") & "</td>"

	' get relevant footnote references 
	
	sql = "select distinct footnote_no "
	sql = sql & "from manager_spell "
	sql = sql & "where manager_id1 = " & rs.Fields("manager_id1")
	if rs.Fields("manager_id2") < 999 then 
		sql = sql & "  and manager_id2 = " & rs.Fields("manager_id2")
	  else 	 	
		sql = sql & "  and manager_id2 is null " 
	end if 
	if not isnull(rs.Fields("spell_no")) then
		if rs.Fields("spell_no") > 0 then 	'zero occurs when spells are to be combined
			sql = sql & "  and spell_no = " & rs.Fields("spell_no")
		end if	
	end if	
	sql = sql & " order by footnote_no "
	
	rs1.open sql,conn,1,2
	
	footnote_no = ""
	Do While Not rs1.EOF
		if not isnull(rs1.Fields("footnote_no")) then 
			footnotes = split(trim(rs1.Fields("footnote_no")),",")	'just in case this spell has more than one footnote, one of which might already have been detected
			for each footnote in footnotes
				if instr(footnote_no,footnote) = 0 then footnote_no = footnote_no & rtrim(rs1.Fields("footnote_no")) & ","
			next
		end if	
		rs1.MoveNext
	Loop
	rs1.close
	
	if footnote_no > "" then footnote_no = left(footnote_no,len(footnote_no)-1)	'remove final comma
	
	outline  = outline & "<td style=""text-align: center"">" & footnote_no & "</td>"
	
	if isnull(rs.Fields("penpiclen")) then
		outline = outline & "<td class=""manager_name "">" & managers & "</td>"
	  else
		outline = outline & "<td id=""cell-a" & rowid & """ class=""manager_name manager_name_hover "">"
		outline = outline & "<img style=""vertical-align: top; margin:0 2px 0 0; padding:0"" src=""images/more.png"">" & managers & "</td>"

		if managers = focusname or managers = focusname & " 1" then 
			dialoghold = "<div id=""dialog"" title=""Manager Focus"" class=""cell-a" & rowid & " manager_name_hover"">"
  			dialoghold = dialoghold & "Click here to focus on<br><span style=""font-size: 14px; font-weight: bold"">" & focusname & "</span><br>or close this box to access all managers" 
			dialoghold = dialoghold & "</div>"
		end if
	end if

	' get relevant from & to dates 
	
	sql = "select from_date as sortdate, convert(varchar,from_date,106) as from_date, convert(varchar,to_date,106) as to_date "
	sql = sql & "from manager_spell "
	sql = sql & "where manager_id1 = " & rs.Fields("manager_id1")
	if rs.Fields("manager_id2") < 999 then 
		sql = sql & "  and manager_id2 = " & rs.Fields("manager_id2")
	  else 	 	
		sql = sql & "  and manager_id2 is null " 
	end if 
	if not isnull(rs.Fields("spell_no")) then
		if rs.Fields("spell_no") > 0 then 	'zero occurs when spells are to be combined
			sql = sql & "  and spell_no = " & rs.Fields("spell_no")
		end if	
	end if	
	sql = sql & " order by sortdate "
	
	rs1.open sql,conn,1,2
		
	fromdates = ""
	todates = ""	

	Do While Not rs1.EOF
		date1 = rs1.Fields("from_date")
		if left(date1,1) = "0" then date1 = mid(date1,2)
		fromdates = fromdates & date1 & "<br>"
		date2 = rs1.Fields("to_date")
		if left(date2,1) = "0" then date2 = mid(date2,2)
		todates = todates & date2 & "<br>"
		rs1.MoveNext
	Loop
	rs1.close
	
	fromdates = left(fromdates,len(fromdates)-4)		' drop the final <br>
	todates = left(todates,len(todates)-4)				' drop the final <br>
	
	outline  = outline & "<td class=""date"">" & fromdates & "</td>"
	outline  = outline & "<td class=""date"">" & todates & "</td>"
	
	outline  = outline & "<td>" & daysic & "</td>"
	
	outline  = outline & "<td style=""border-top:0; border-bottom:0;"">&nbsp;</td>"
	outline  = outline & "<td>" & rs.Fields("P") & "</td>"	
	outline  = outline & "<td>" & rs.Fields("W") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("D") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("L") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("F") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("A") & "</td>" 
	outline  = outline & "<td>" & rs.Fields("PO") & "</td>" 

	outline  = outline & "<td>" & Rightpad1(Cstr(rs.Fields("Wperc"))) & "</td>"  
	outline  = outline & "<td>" & Rightpad1(Cstr(rs.Fields("Dperc"))) & "</td>"  
	outline  = outline & "<td>" & Rightpad1(Cstr(rs.Fields("Lperc"))) & "</td>"  
	outline  = outline & "<td class=""leftpad"">" & Rightpad2(Cstr(rs.Fields("Favg"))) & "</td>"  
	outline  = outline & "<td class=""leftpad"">" & Rightpad2(Cstr(rs.Fields("Aavg"))) & "</td>" 
	outline  = outline & "<td class=""leftpad"">" & Rightpad2(Cstr(rs.Fields("POavg"))) & "</td>"

	outline  = outline & "<td style=""border-top:0; border-bottom:0;"">&nbsp;</td>"
	outline  = outline & "<td>" & rs.Fields("CP") & "</td>"
	outline  = outline & "<td>" & rs.Fields("CW") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("CD") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("CL") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("CF") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("CA") & "</td>"
	outline  = outline & "<td>" & Rightpad1(Cstr(rs.Fields("CWperc"))) & "</td>"  
	outline  = outline & "</tr>" 
	
	outline  = outline & "<tr id=""row-b" & rowid & """ class=""hide""><td id=""cell-b" & rowid & """  class=""green"" style=""padding:0 18px 12px;"" colspan=""28""></td></tr>"
	n = n + 1
	
	rs.MoveNext
	
Loop

response.write(outline)
	
rs.close
	
%>

    <tr class="footnote">
    <td style="text-align:left; padding:15px 10px;" colspan="28"><span style="font-weight: bold; margin:0 22px">Footnotes:</span>
    <ol>
    
<%
sql = "select footnote_no, footnote "
sql = sql & "from manager_footnote "
sql = sql & "order by footnote_no "

rs.open sql,conn,1,2 

Do While Not rs.EOF

	response.write("<li style=""margin:6px 0;"">" & rs.Fields("footnote") & "</li>")
		
	rs.MoveNext
	
Loop
rs.close
conn.close
%>
    
    </ol>
    </td>
    </tr>
	</table>
	
<%
Function Rightpad1(number)
	if instr(number,".") = 0 then 
		Rightpad1 = number & ".0"
	  else	
	  	Rightpad1 = number
	end if	
End Function
Function Rightpad2(number)
	if instr(number,".") = 0 then 
		Rightpad2 = number & ".00"
	  elseif instrRev(number,".") = len(number)-1 then 
		Rightpad2 = number & "0"
	  else 
	  	Rightpad2 = number
	end if	
End Function		 
%>


</div>

<div id="overlaydiv"></div>"
<%
response.write(dialoghold)
%>
	
<!--#include file="base_code.htm"-->
</body></html>