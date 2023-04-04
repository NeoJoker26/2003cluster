<style>
<!--
div#table1 tr td { border: 0px none; }
div#table1 td.a {border-left: 1px dotted #c0c0c0; border-right: 1px dotted #c0c0c0; }
div#table1 td.b {border-bottom-style: none; border-right-style: none;}
div#table1 td.c {border-bottom-style: none; border-left-style: none; border-right-style: none;}
div#table1 td.d {border-bottom-style: none; border-left-style: none; border-right: 1px dotted #c0c0c0; }
div#table1 td.t {border-top: 1px solid #c0c0c0; }
div#table1 td.head {border: 0px solid #c0c0c0; padding-bottom: 6px; font-size: 11px; font-weight: bold; color:#006e32; }

-->
   </style>
<script language="javascript">
function HeadToggle(item) {
   obj=document.getElementById(item);
   visible=(obj.style.display!="none")
   key=document.getElementById("x" + item);
   if (visible) {
     obj.style.display="none";
     key.innerHTML="[+]";
   } else {
      obj.style.display="block";
      key.innerHTML="[-]";
   }
}

function Toggle(item,clickon) {
   obj=document.getElementById(item);
   objdate=document.getElementById("d" + item);
   GetDetails(objdate.innerHTML);

   visible=(obj.style.display!="none")
   key=document.getElementById("x" + item);
   if (visible) {
     obj.style.display="none";
     key.innerHTML= '[+<span style="font-family:verdana;">' + clickon + '</span>]';
   } else {
      obj.style.display="block";
      key.innerHTML='[-<span style="font-family:verdana;">' + clickon + '</span>]';
   }
}

var xmlHttp

function GetDetails(str)
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
</script>
<%
Dim conn,sql,rs,rslineup,rsgoals, displaydate, years, yearslast, opposition, outline, outline1, outline2, outlinematch, called, fullteam, crowd, venue, work1, tagno
Dim latestdates(5,1), counts(8,2),i,j, tableview, heading1, heading2, headtext, competition, check1, check2, check3, check4, clickon, topclass, season_years1, season_years2, season_text, restrictions

Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rsgoals = Server.CreateObject("ADODB.Recordset")
Set rslineup = Server.CreateObject("ADODB.Recordset")

%><!--#include file="conn_read.inc"--><%
%>