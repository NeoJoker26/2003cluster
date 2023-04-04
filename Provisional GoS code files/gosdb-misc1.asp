<%@ Language=VBScript %> 
<% Option Explicit %>
<% dim scope
scope = Request.Form("scope")
if scope = "" then scope = Request.Querystring("scope")	'try for a url parameter
scope = replace(scope," ","")
%>
<!DOCTYPE html PUBLIC "-//w3c//dtd html 4.0 transitional//en">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="Author" content="Trevor Scallan">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<title>GoS-DB Miscellaneous Report</title>
<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
#sumtable td {border: 1px solid #c0c0c0; text-align:right; margin: 0; white-space:nowrap; padding-left:1; padding-right:2; padding-top:1; padding-bottom:1}
#sumtable .l {padding-left: 4; border-left: 2px solid #c0c0c0 ;}
#sumtable .r {padding-right: 6; border-right: 2px solid #c0c0c0 ;}
#sumtable .rowhlt { background-color: #d5e9d7; }
-->
</style>

</head>

<body>

<!--#include file="top_code.htm"-->
<%
Dim conn,sql,rs, outline 


Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%
%>
  <center>
  <table border="0" cellspacing="0" style="border-collapse: collapse" 
  cellpadding="0" width="980">
    <tr>
    <td width="260" valign="top" style="text-align:center;">

	<p style="text-align: center; margin-top:0; margin-bottom:3">
	<a href="gosdb.asp"><font color="#404040"><img border="0" src="images/gosdb-small.jpg" align="left"></font></a><font color="#404040"> 
	<b><font style="font-size: 15px">Search by<br>
	</font></b><span style="font-size: 15px"><b>Player</b></span></font><p style="text-align: center; margin-top:0; margin-bottom:0">
	<b>
	<a href="gosdb.asp">Back to<br>GoS-DB Hub</a></b></p>

	</td>
    
  	<td width="460" align="center" style="text-align: center" valign="top">	
	<p style="margin-top:12; margin-bottom:0; text-align:center; font-size:18px; color:#006E32">
    MISCELLANEOUS REPORTS</p>  
    
	<p style="margin-top:6; margin-bottom:0; text-align:center; font-size:13px">
    <b>Report 1:  Competition Totals</b></p>  
    </td>
        
	<td width="260" valign="top"  align="justify">
	'<span style="font-size: 10px">Miscellaneous Reports' is an ever-growing collection of pages that reflect 
    broad aspects of Argyle's playing history. If you have an idea for another, 
    please get in touch. </span>
     
    </td>
    </tr>
	</table>
      
<%
 outline = ""
sql = "WITH CTE1 AS ( "
sql = sql & "select case LFC when 'F' then 'League' when 'L' then 'Non-league' else 'Cup' end as comptype1, "
sql = sql & "case LFC when 'F' then 'League' + ' tier ' + cast(tier as varchar) when 'L' then 'Non-league' else 'Cup' end as comptype2, "
sql = sql & "competition, "
sql = sql & "1 as p, "
sql = sql & "case when goalsfor > goalsagainst then 1 else 0 end as w, "
sql = sql & "case when goalsfor = goalsagainst then 1 else 0 end as d, "
sql = sql & "case when goalsfor < goalsagainst then 1 else 0 end as l, "
sql = sql & "goalsfor as f, goalsagainst as a, attendance at,  "
sql = sql & "case when homeaway = 'H' then 1 else 0 end as hp, "
sql = sql & "case when homeaway = 'H' and goalsfor > goalsagainst then 1 else 0 end as hw, "
sql = sql & "case when homeaway = 'H' and goalsfor = goalsagainst then 1 else 0 end as hd, "
sql = sql & "case when homeaway = 'H' and goalsfor < goalsagainst then 1 else 0 end as hl, "
sql = sql & "case when homeaway = 'H' then goalsfor else 0 end as hf, "
sql = sql & "case when homeaway = 'H' then goalsagainst else 0 end as ha, "
sql = sql & "case when homeaway = 'H' then attendance else NULL end as hat, "
sql = sql & "case when homeaway <> 'H' then 1 else 0 end as ap, "
sql = sql & "case when homeaway <> 'H' and goalsfor > goalsagainst then 1 else 0 end as aw, "
sql = sql & "case when homeaway <> 'H' and goalsfor = goalsagainst then 1 else 0 end as ad, "
sql = sql & "case when homeaway <> 'H' and goalsfor < goalsagainst then 1 else 0 end as al, "
sql = sql & "case when homeaway <> 'H' then goalsfor else 0 end as af, "
sql = sql & "case when homeaway <> 'H' then goalsagainst else 0 end as aa, "
sql = sql & "case when homeaway <> 'H' then attendance else NULL end as aat "
sql = sql & "from v_match_all join season on date between date_start and date_end "
sql = sql & ") "
sql = sql & "select case when grouping(comptype1) = 1 then 'zzz' else comptype1 end as comptype1, "
sql = sql & " case when grouping(comptype2) = 1 then 'zzzz' else comptype2 end as comptype2, "
sql = sql & " case when grouping(competition) = 1 then 'zzzzz' else competition end as competition, "
sql = sql & "sum(p) as P, sum(w) as W, sum(d) as D, sum(l) as L, sum(f) as F, sum(a) as A, avg(at) as AT, "
sql = sql & "sum(hp) as HP, sum(hw) as HW, sum(hd) as HD, sum(hl) as HL, sum(hf) as HF, sum(ha) as HA, avg(hat) as HAT, "
sql = sql & "sum(ap) as AP, sum(aw) as AW, sum(ad) as AD, sum(al) as AL, sum(af) as AF, sum(aa) as AA, avg(aat) as AAT  "
sql = sql & "from CTE1 "  
sql = sql & "group by comptype1, comptype2, competition with rollup "
sql = sql & "order by comptype1, comptype2, competition "
rs.open sql,conn,1,2
	
outline = outline & "<table id=""sumtable"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""margin-top: 9; border-collapse: collapse"">"
 	outline  = outline & "<tr>"
      outline  = outline & "<td style=""border: 0px none white;"" colspan=""2"">&nbsp;</td>"
      outline  = outline & "<td colspan=""6"" class=""l r"" style=""text-align:center; border-top-color: #C0C0C0; border-top-width: 1""><b>Home</b></td>"
      outline  = outline & "<td colspan=""6"" class=""l r"" style=""text-align:center; border-top-color: #C0C0C0; border-top-width: 1""><b>Away or Neutral</b></td>"
      outline  = outline & "<td colspan=""6"" class=""l r"" style=""text-align:center; border-top-color: #C0C0C0; border-top-width: 1""><b>Totals</b></td>"
      outline  = outline & "<td colspan=""3"" class=""l r"" style=""text-align:center; border-top-color: #C0C0C0; border-top-width: 1""><b>Attendance</b></td>"
    outline  = outline & "</tr>"
    outline  = outline & "<tr>"
      outline  = outline & "<td style=""padding:0 8 0 8; text-align: left""><b>Type</b></td>"
      outline  = outline & "<td style=""padding:0 8 0 8; text-align: left""><b>Competition</b></td>"
      outline  = outline & "<td class=""l""><b>P</b></td>"
      outline  = outline & "<td><b>W</b></td>"
      outline  = outline & "<td><b>D</b></td>"
      outline  = outline & "<td><b>L</b></td>"
      outline  = outline & "<td><b>F</b></td>"
      outline  = outline & "<td class=""r""><b>A</b></td>"
      outline  = outline & "<td class=""l""><b>P</b></td>"
      outline  = outline & "<td><b>W</b></td>"
      outline  = outline & "<td><b>D</b></td>"
      outline  = outline & "<td><b>L</b></td>"
      outline  = outline & "<td><b>F</b></td>"
      outline  = outline & "<td class=""r""><b>A</b></td>"
      outline  = outline & "<td class=""l""><b>P</b></td>"
      outline  = outline & "<td><b>W</b></td>"
      outline  = outline & "<td><b>D</b></td>"
      outline  = outline & "<td><b>L</b></td>"
      outline  = outline & "<td><b>F</b></td>"
      outline  = outline & "<td class=""r""><b>A</b></td>"
      outline  = outline & "<td class=""l"" style=""text-align: center""><b>Avg<br>Home</b></td>"
      outline  = outline & "<td class=""r"" style=""text-align: center""><b>Avg<br>Away</b></td>"
    outline  = outline & "</tr>"

Do While Not rs.EOF

  if rs.Fields("comptype1") <> "zzz" and rs.Fields("comptype1") <> "League" and rs.Fields("comptype2") = "zzzz" and rs.Fields("competition") = "zzzzz" then 
	'don't need this sub-total line
   else
		
	if rs.Fields("comptype1") = "zzz" and rs.Fields("comptype2") = "zzzz" and rs.Fields("competition") = "zzzzz" then
		outline  = outline & "<tr style=""font-weight:bold; color:#004b18;"" onmouseover=""this.className = 'rowhlt';"" onmouseout=""this.className = '';""><td style=""padding:0 8 0 8; text-align: left"" colspan=""2"">Grand Totals/Averages</td>"
	  elseif rs.Fields("comptype1") = "League" and rs.Fields("comptype2") = "zzzz" and rs.Fields("competition") = "zzzzz" then
		outline  = outline & "<tr style=""color:#004b18;"" onmouseover=""this.className = 'rowhlt';"" onmouseout=""this.className = '';""><td style=""padding:0 8 0 8; text-align: left"" colspan=""2"">Football League Totals/Averages</td>"
	  elseif rs.Fields("competition") = "zzzzz" then
		outline  = outline & "<tr style=""color:#004b18;"" onmouseover=""this.className = 'rowhlt';"" onmouseout=""this.className = '';""><td style=""padding:0 8 0 8; text-align: left"" colspan=""2"">" & rs.Fields("comptype2") & " Totals/Averages</td>"
	  else
		outline  = outline & "<tr onmouseover=""this.className = 'rowhlt';"" onmouseout=""this.className = '';""><td style=""padding:0 8 0 8; text-align: left"">" & rs.Fields("comptype2") & "</td>"
		outline  = outline & "<td style=""padding:0 8 0 8; text-align: left"">" & rs.Fields("competition") & "</td>"
	end if
	
	outline  = outline & "<td class=""l"">" & rs.Fields("HP") & "</td>"
	outline  = outline & "<td>" & rs.Fields("HW") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("HD") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("HL") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("HF") & "</td>"  
	outline  = outline & "<td class=""r"">" & rs.Fields("HA") & "</td>" 
	outline  = outline & "<td class=""l"">" & rs.Fields("AP") & "</td>"
	outline  = outline & "<td>" & rs.Fields("AW") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("AD") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("AL") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("AF") & "</td>"  
	outline  = outline & "<td class=""r"">" & rs.Fields("AA") & "</td>"
	outline  = outline & "<td class=""l"">" & rs.Fields("P") & "</td>"	
	outline  = outline & "<td>" & rs.Fields("W") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("D") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("L") & "</td>"  
	outline  = outline & "<td>" & rs.Fields("F") & "</td>"  
	outline  = outline & "<td class=""r"">" & rs.Fields("A") & "</td>"  
	outline  = outline & "<td class=""l"" align=""right"">" & rs.Fields("HAT") & "</td>"
	outline  = outline & "<td class=""r"" align=""right"">" & rs.Fields("AAT") & "</td>"
	outline  = outline & "</tr>" 
  end if	  
  rs.MoveNext
Loop
	
rs.close

outline = outline & "</table>"
	
response.write(outline)

conn.close
%>	
	
<br>

<!--#include file="base_code.htm"-->
</body>

</html>