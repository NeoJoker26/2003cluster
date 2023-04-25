
<%@ Language=VBScript %>
<% Option Explicit %>

<html>
<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<title>Greens on Screen Database</title>

<base target="_self">
<link rel="stylesheet" type="text/css" href="gos2.css">
<style>
<!--

td {text-align:left; margin: 0; padding: 3px 2px 3px 3px;  font-size:11px;} 
p {text-align:left; margin: 0; padding: 0 2px 6px 3px;  font-size:11px;} 

#gottable1 {border-collapse: collapse; border: 1px solid #c0c0c0; width:280px; }
#gottable1 td {text-align:left; margin: 0; padding: 2px 4px; border: 1px solid #c0c0c0; font-family: "Trebuchet MS",helvetica,verdana,arial,sans-serif; font-size: 11px; }
#gottable1 p {margin: 3px 0; padding: 0; font-family: verdana,arial,sans-serif; font-size: 10px; font-weight:bold; } 
#gottable1 .right {text-align: right; } 
#gottable1 .tah {font-family: "Trebuchet MS",helvetica,verdana,arial,sans-serif; } 

#gottable2 {border-collapse: collapse; border: 1px solid #c0c0c0; width:280px; }
#gottable2 td {text-align:left; margin: 0; padding: 2px 4px; border: 1px solid #c0c0c0; font-family: "Trebuchet MS",helvetica,verdana,arial,sans-serif; font-size: 11px; } 
#gottable2 p {margin: 3px 0; padding: 0; font-family: verdana,arial,sans-serif; font-size: 10px; font-weight:bold; } 
#gottable2 .right {text-align: right; }
#gottable2 .tah {font-family: "Trebuchet MS",helvetica,verdana,arial,sans-serif; } 

.button1 {
    padding: 3px 8px 3px; 
    margin: 0 2px 6px 0;
    font-family: verdana, sans-serif;
    display: inline-block;
    white-space: nowrap;
    font-size:11px;
    position:relative;
    outline: none;
    overflow: visible;
    cursor: pointer;
    border-radius: 2px;
    border: 1px solid #808080;
    color: #000000 !important; 
    background: linear-gradient(#e0f0e0,#d0e0d0);  
}

.button1:hover {
    background: linear-gradient(#c8e0c7,#c6dbc5);
    border: 1px solid #000000; 
}

#misctable { 
	border:0px none;
	border-collapse: collapse; 
	margin-left:44px; 
	margin-right:0; 
	margin-top:0; 
	margin-bottom:12px;
}
#misctable td { 
	padding: 2px 3px; 
}

-->

</style>


<script language="javascript">

function GetTable(side,comp) { 

try { 
        // Moz supports XMLHttpRequest. IE uses ActiveX. 
        // browser detction is bad. object detection works for any browser 
        xmlhttp = window.XMLHttpRequest?new XMLHttpRequest(): new ActiveXObject("Microsoft.XMLHTTP"); 
} catch (e) { 
        // browser doesn't support ajax. handle however you want 
          alert ("Sorry, your browser does not support this function.");
  		  return;
} 


// the xmlhttp object triggers an event everytime the status changes 
// triggered() function handles the events 
xmlhttp.onreadystatechange = triggered; 

// open takes in the HTTP method and url.
//document.body.style.cursor='wait';        
var url="gosdb-getmenupagedetails.asp";
url=url+"?side="+side;
url=url+"&comp="+comp;
url=url+"&sid="+Math.random();

xmlhttp.open("GET", url, true); 
 
// send the request. if this is a POST request we would have 
// sent post variables: send("name=aleem&gender=male) 
// Moz is fine with just send(); but 
// IE expects a value here, hence we do send(null); 
xmlhttp.send(null);
//document.body.style.cursor='auto';  
} 
 
function triggered() { 
// if the readyState code is 4 (Completed) 
// and http status is 200 (OK) we go ahead and get the responseText 
// other readyState codes: 
// 0=Uninitialised 1=Loading 2=Loaded 3=Interactive 
if (xmlhttp.readyState == 4) { 
        // xmlhttp.responseText object contains the response.
        var textsplit = xmlhttp.responseText.split("^");
        if (textsplit[0] == "left" || textsplit[0] == "both") { document.getElementById('left').innerHTML = textsplit[1]; } 
        if (textsplit[0] == "right" || textsplit[0] == "both") { document.getElementById('right').innerHTML = textsplit[2]; } 
} 
} 

</script>

</head>

<body onLoad="javascript:GetTable('both','all');">
<!--#include file="top_code.htm"-->

<%
Dim conn,sql,rs,lastdate

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
%><!--#include file="conn_read.inc"--><%

sql = "select max(date) as lastdate "
sql = sql & "from match " 

rs.open sql,conn,1,2
lastdate = WeekdayName(Weekday(rs.Fields("lastdate"))) & " " & Day(rs.Fields("lastdate")) & " " & MonthName(Month(rs.Fields("lastdate"))) & ", " & Year(rs.Fields("lastdate"))
rs.close 
conn.close
%>
<center>

<table border="0" cellpadding="0" cellspacing="0" 
style="border-collapse: collapse; margin-top:20px 0;" bordercolor="#111111" 
width="984">
  <tr>
    <td id="cellleft" valign="top" width="320">
    <p style="margin-top: 0; margin-bottom: 9px; margin-left:0; margin-right:15px; ">
    Some of the information in GoS-DB differs from other records. Where 
    anomalies have been 
    found, the data has been verified using original newspaper reports to 
    ensure a very high degree of accuracy.</p>
    
	<div id="left" align="left" style="width:320px">
    Retrieving data <img border="0" src="images/ajax-loader.gif" align="texttop">
    <img border="0" src="images/dummy.gif" height="220px" width="1px" align="texttop">
    </div>
    <div id="left1" style="padding-right: 45px;">
    <p style="margin-top: 18"><font color="#004438">
    <span style="font-size: 12px"><b>First time here</b></span></font><font color="#004438" style="font-size: 12px"><b>?</b></font></p> 
    <p>GoS-DB provides a vast array of facts and figures from Argyle's history. 
    If this is your first 
    visit, the <b><u><a href="gosdb-about.asp"><u>Getting Started</u></a></u></b> 
    page is a useful place to begin.</p> 
    </div>
	
    </td>

    <td valign="top" align="center" style="text-align: center" width="344">
    <map name="map1">
    <area title="Septimus Atterbury 1907-1915, 1919-21" href="#" shape="polygon" 
    coords="32, 54, 28, 65, 31, 77, 33, 86, 18, 98, 10, 110, 8, 126, 5, 135, 82, 135, 82, 122, 76, 107, 64, 100, 55, 92, 51, 81, 56, 71, 54, 56, 48, 52, 36, 50">
    <area title="Sammy Black 1924-1938" href="#" shape="polygon" 
    coords="83, 135, 83, 124, 82, 115, 78, 107, 72, 103, 75, 98, 84, 95, 87, 93, 87, 86, 86, 80, 86, 70, 88, 63, 93, 59, 98, 57, 107, 59, 109, 66, 109, 72, 109, 78, 109, 80, 108, 88, 112, 92, 118, 95, 125, 99, 120, 104, 119, 112, 121, 121, 123, 130, 123, 135">
    <area title="Wilf Carter 1957-1964" href="#" shape="polygon" 
    coords="127, 135, 127, 128, 124, 118, 123, 115, 123, 109, 128, 103, 136, 100, 145, 97, 148, 92, 148, 87, 147, 80, 146, 72, 146, 66, 148, 59, 152, 56, 161, 53, 166, 58, 171, 69, 166, 76, 163, 81, 162, 86, 162, 88, 174, 91, 182, 96, 186, 104, 187, 113, 187, 123, 189, 133, 130, 135">
    <area title="Paul Mariner 1973-1976" href="#" shape="polygon" 
    coords="193, 133, 192, 121, 190, 110, 186, 101, 183, 97, 183, 94, 194, 89, 196, 84, 194, 80, 193, 70, 194, 62, 196, 57, 202, 53, 209, 52, 216, 54, 219, 60, 222, 75, 220, 80, 224, 83, 234, 87, 237, 99, 228, 107, 228, 114, 226, 120, 225, 129, 225, 135, 196, 135">
    <area title="Mickey Evans 1990-1997, 2001-2006" href="#" 
    shape="polygon" 
    coords="226, 134, 227, 120, 230, 110, 234, 102, 243, 97, 250, 93, 257, 88, 258, 81, 256, 76, 253, 67, 253, 60, 259, 54, 266, 53, 273, 54, 276, 58, 278, 66, 276, 75, 273, 80, 273, 86, 277, 91, 286, 94, 291, 97, 300, 104, 305, 116, 306, 127, 306, 135">
    </map>
    <img border="0" src="images/gosdb.jpg" usemap="#map1">
    <p style="margin: 2px 0; text-align:center; font-size: 10px;">Greens on Screen's history pages are underpinned by GoS-DB, 
    a relational database of over 100,000 facts and stats from 1903 to <%response.write(lastdate)%>.
 	</p>
 	
 	<div id="menutab" style="margin: 0 auto 12px">
 	<p style="text-align:center; font-weight:bold; margin: 6px 0 4px;">Search by ...</p>
	<a class="button1" href="gosdb-match.asp">MATCH DATE</a>
	<a class="button1" href="gosdb-dates.asp">DAY OF YEAR</a>
    <a class="button1" href="gosdb-seasons.asp">SEASON</a>
    <a class="button1" href="gosdb-players0.asp">PLAYER</a>
    <a class="button1" href="gosdb-managers.asp">MANAGER</a>
    <a class="button1" href="gosdb-headtohead.asp">OPPOSITION</a>
	</div>
 
 
    <p style="text-align:center; font-weight:bold; padding:0; margin-bottom:5px;">or select one of these reports ...</p>
    
    <table id="misctable">
      <tr>
        <td style="text-align: right">1</td>
        <td><a href="gosdb-misc1.asp">Competition Totals</a></td>
      </tr>
      <tr>
        <td style="text-align: right">2</td>
        <td><a href="gosdb-misc2.asp">Consecutive Results</a></td>
      </tr>
      <tr>
        <td style="text-align: right">3</td>
        <td><a href="gosdb-misc3.asp">Attendance Highs and Lows</a></td>
      </tr>
      <tr>
        <td style="text-align: right">4</td>
        <td><a href="gosdb-misc4.asp">Top Substitutes</a></td>
      </tr>
      <tr>
        <td style="text-align: right">5</td>
        <td><a href="gosdb-misc5.asp">Youngest and Oldest</a></td>
      </tr>
      <tr>
        <td style="text-align: right">6</td>
        <td><a href="gosdb-misc6.asp">Best and Worst Starts</a></td>
      </tr>
      <tr>
        <td style="text-align: right">7</td>
        <td><a href="gosdb-misc7.asp">Football League by Decade</a></td>
      </tr>
      <tr>
        <td style="text-align: right">8</td>
        <td><a href="gosdb-misc8.asp">Football League by Calendar Year</a></td>
      </tr>
      <tr>
        <td style="text-align: right">9</td>
        <td><a href="gosdb-misc9.asp">Success Rankings by Opposition</a></td>
      </tr>
      <tr>
        <td style="text-align: right">10</td>
        <td><a href="gosdb-misc10.asp">Goalscorers by Season</a></td>
      </tr>
      <tr>
        <td style="text-align: right">11</td>
        <td><a href="gosdb-misc11.asp">Score Counts</a></td>
      </tr>
      <tr>
        <td style="text-align: right">12</td>
        <td><a href="gosdb-misc12.asp">Consecutive Appearances</a></td>
      </tr>
    </table>


 	<p style="text-align:center; margin-left:15px; margin-right:15px; margin-top:12px">Any ideas? If you can think of a useful addition<br>
    to this list, please use the Contact Us link.</p> 

    </td>
            
    <td valign="top" align="right" width="320" style="text-align: right">
    <p style="text-align: right; margin-right: 0; margin-bottom:9; margin-left:12">
    We are very confident that GoS-DB is as least as accurate as  
    other Argyle-related resources, and in many cases   
    corrects long-standing mistakes. Please use 'Contact 
    Us' if you can make it even better.</p>

	<div id="right" align="right" style="width:320px">
	<p style="text-align: right">Retrieving data <img border="0" src="images/ajax-loader.gif" align="texttop">
	<img border="0" src="images/dummy.gif" height="220px" width="1px" align="texttop">
	</p>
	</div>
	<div id="right1" style="padding-left: 0px; width:340px;">
	<p style="margin: 18px 0 -1px 36px;"><font color="#004438" style="font-size: 12px"><b>GoS-DB History</b></font></p> 
	<iframe style="padding:0; margin:0; border:0px none; height:180px" name="I1" frameborder="0" src="gosdb-history.htm">
    Your browser does not support inline frames or is currently configured not to display inline frames.</iframe>

    </div>

    </td>
  </tr>
  </table>
</td></tr>
</table>
<!--#include file="base_code.htm"-->
</body>
</html>