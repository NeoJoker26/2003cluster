<%@ Language=VBScript %>
<% Option Explicit %>

<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Greens on Screen</title>

<link rel="stylesheet" type="text/css" href="gos2.css">

<style>
<!--
td {font-size: 11px;}
-->
</style>
  
</head>
  
<body><!--#include file="top_code.htm"-->

<center>
<form action="photocaptions1.asp" method="post" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1">
<center>

<table border="0" cellspacing="5" style="margin-bottom:12px; border-collapse: collapse; text-align:center" bordercolor="#111111" width="374">
  <tr>
    <td style="text-align: left" colspan="2" width="354">
<p style="text-align: center; margin-bottom: 18; margin-top:12">
<font color="#47784D" style="font-size: 18px"><b>Photo Captions</b></font></p>
    </td>
  </tr>
  <tr>
    <td style="text-align: left" width="75">
    <p style="margin-top: 0; margin-bottom: 0">Set Code: </td>
    <td style="text-align: left" width="264">
	<p style="margin-top: 0; margin-bottom: 0">
    <input type="text" name="code" size="14"</td>
  </tr>
  <tr>
    <td style="text-align: left" width="75">
    <p style="margin-top: 0; margin-bottom: 0">Set Type: </td>
    <td style="text-align: left" width="264">
	<p style="margin-top: 0; margin-bottom: 0">
    <select size="1" name="type">
    <option selected value="M">Match</option>
    <option value="O">Other Match</option>
    <option value="F">Pre-season</option>
    <option value="E">Random Event</option>
	<option value="H">Redevelopment</option>
	<option value="S">Season</option>
	<option value="W">Slideshow</option>
  </tr>
  </tr>
  <tr>
    <td style="text-align: left" width="75">
    <p style="margin-top: 0; margin-bottom: 0">Process: </td>
    <td style="text-align: left" width="264">
	<p style="margin-top: 0; margin-bottom: 0">
    <select size="1" name="process">
    <option selected value="1">1. Normal</option>
    <option value="2">2. Changed photos</option>
    <option value="3">3. Destroy & start again</option>
    </select>
  </tr>
  <tr>
    <td style="text-align: left" colspan="2" width="354">
	<p style="margin-top: 12">
	<input type="submit" name="b1" value="Process uploaded images ready for captions" style="width: 260px; font-size: 12px; margin-left:0; margin-right:0; padding:0;">&nbsp;&nbsp; <b>
	<font color="#808080">|</font>&nbsp; <a target="_top" href="index.asp">Cancel</a></b></td>
  </tr>
  </table>
</form>

<table border="0" cellspacing="5" style="margin-left:12px; border-collapse: collapse; text-align:left" bordercolor="#111111" width="441">
  <tr>
    <td style="text-align: left" width="397" valign="top" colspan="2">
	<p style="margin-top: 6; margin-bottom: 3"><b>Process Notes</b></td>
    </tr>
  <tr>
    <td style="text-align: left" width="3" valign="top">
	1:</td>
    <td style="text-align: left" width="394">
	Normal: captions are being written for the first time or are being 
    amended (also use when your normal practice is to add a few captions at a 
    time).</td>
  </tr>
  <tr>
    <td style="text-align: left" width="3" valign="top">
	2:</td>
    <td style="text-align: left" width="394">
	Changed photos: photos to be rebuilt from a refreshed upload; for example, 
    when 
    individual images have been altered (e.g. cropped), new photos have been 
    added or photos have been dropped from the set. Existing captions 
    will be preserved whenever possible, but please check.
	</td>
  </tr>
  <tr>
    <td style="text-align: left" width="3" valign="top">
	3: 
	</td>
    <td style="text-align: left" width="394">
	Destroy &amp; start again: all existing photos and all captions will be 
    deleted, and the photo set will be treated as 'first time'. This is a 
    desperate measure, the 'nuclear option'; <i>not to be used lightly!&nbsp;&nbsp;
    </i>
	</td>
  </tr>
  </table>


</center>
<!--#include file="base_code.htm"-->
</body>

</html>